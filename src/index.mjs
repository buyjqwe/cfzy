/**
 * Cloudflare Worker - 自动化作业批改
 *
 * 功能:
 * 1. GET 请求: 返回一个 HTML 上传页面。
 * 2. POST 请求 (multipart/form-data):
 * - 立即响应 "上传成功"，防止浏览器超时。
 * - 在后台 (ctx.waitUntil) 执行所有耗时操作：
 * a. 获取 Microsoft Graph API 令牌 (客户端凭据流)。
 * b. 将作业 .zip 包上传到 OneDrive。
 * c. 解压主 .zip (unzipit)。
 * d. 遍历每个学生的 .zip (unzipit)。
 * e. 读取学生文件，构建 Gemini API 请求。
 * f. 调用 Gemini API 进行批改。
 * g. 将批改结果 (report.json) 上传回该学生的 OneDrive 文件夹。
 *
 * 环境变量 (Secrets) - 必须在 Cloudflare 仪表盘或 deploy.yml 中设置:
 * - GOOGLE_API_KEY: Gemini API 密钥。
 * - MS_CLIENT_ID: Microsoft Azure App Client ID。
 * - MS_CLIENT_SECRET: Microsoft Azure App Client Secret。
 * - MS_TENANT_ID: Microsoft Azure Tenant ID (租户ID)。
 * - MS_USER_ID: 目标 OneDrive 账户的 Microsoft Graph User ID (对象ID)。
 */

import { unzip } from 'unzipit'; // 用于解压 .zip 文件
import { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } from '@google/generative-ai'; // Gemini API

// --- 全局常量 ---
const ONEDRIVE_BASE_PATH = 'root:/Apps/HomeworkGrader'; // OneDrive 上的基础存储路径

// Gemini API 安全设置 (设为最低，允许处理各种内容)
const GEMINI_SAFETY_SETTINGS = [
  { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
];

export default {
  /**
   * Worker 入口点
   * @param {Request} request - 传入的请求
   * @param {object} env - 环境变量 (Secrets)
   * @param {object} ctx - 执行上下文 (用于 ctx.waitUntil)
   * @returns {Promise<Response>}
   */
  async fetch(request, env, ctx) {
    // 检查所有必需的密钥是否存在
    const requiredKeys = ['GOOGLE_API_KEY', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_TENANT_ID', 'MS_USER_ID'];
    for (const key of requiredKeys) {
      if (!env[key]) {
        console.error(`[配置错误] 环境变量 ${key} 未设置。`);
        return new Response(`服务器配置错误: ${key} 未设置。`, { status: 500 });
      }
    }

    try {
      if (request.method === 'GET') {
        // GET 请求：返回 HTML 上传页面
        return new Response(getHtmlPage(), {
          status: 200,
          headers: { 'Content-Type': 'text/html; charset=utf-8' },
        });
      } else if (request.method === 'POST') {
        // POST 请求：处理文件上传
        const formData = await request.formData();
        const file = formData.get('homeworkZip'); // 'homeworkZip' 必须与 HTML 表单中的 name 匹配

        if (!file || typeof file.name === 'undefined') {
          return new Response('未找到名为 "homeworkZip" 的文件。', { status: 400 });
        }

        if (!file.name.endsWith('.zip')) {
          return new Response('无效的文件类型。请上传 .zip 文件。', { status: 400 });
        }

        // 将文件内容读取为 ArrayBuffer
        const fileBuffer = await file.arrayBuffer();
        const homeworkZipName = file.name; // 例如 "Week5_Homework.zip"

        // 关键：立即响应浏览器，防止超时
        // 同时将真正的处理任务交给 ctx.waitUntil 在后台异步执行
        ctx.waitUntil(
          handleHomeworkProcessing(env, fileBuffer, homeworkZipName)
            .catch(err => {
              // 捕获后台任务的严重错误并记录
              // 注意：这不会发送给用户，因为响应已经发出
              console.error(`[后台处理失败] ${homeworkZipName}: ${err.message}`, err.stack);
              // 可以在这里触发一个错误通知（例如发送到另一个 Worker 或 Webhook）
            })
        );

        // 立即返回成功信息给用户
        return new Response(
          `文件 "${homeworkZipName}" 已收到，正在后台处理批改。这可能需要几分钟时间。`,
          { status: 202 } // 202 Accepted (已接受，正在处理)
        );

      } else {
        // 其他方法 (PUT, DELETE 等)
        return new Response('无效的请求方法。请使用 GET 访问页面，或使用 POST 上传文件。', {
          status: 405, // Method Not Allowed
          headers: { 'Allow': 'GET, POST' },
        });
      }
    } catch (err) {
      console.error(`[请求处理失败] ${err.message}`, err.stack);
      return new Response(`服务器内部错误: ${err.message}`, { status: 500 });
    }
  },
};

/**
 * [后台任务] 异步处理作业的主逻辑
 * @param {object} env - 环境变量
 * @param {ArrayBuffer} fileBuffer - 主 .zip 文件的内容
 * @param {string} homeworkZipName - 主 .zip 文件的名称
 */
async function handleHomeworkProcessing(env, fileBuffer, homeworkZipName) {
  console.log(`[开始处理] ${homeworkZipName}`);

  // 1. 获取 MS Graph API 令牌
  const accessToken = await getMsGraphToken(env);
  if (!accessToken) {
    console.error(`[MS Token 失败] 无法获取 ${homeworkZipName} 的 MS Graph Token。`);
    return; // 无法继续
  }
  console.log(`[MS Token 成功] ${homeworkZipName}`);

  // 2. 将原始 .zip 包上传到 OneDrive (备份)
  // 我们将 homeworkZipName 去掉 .zip 后缀，作为主文件夹名
  const baseFolderName = homeworkZipName.replace(/\.zip$/i, '');
  const backupPath = `${ONEDRIVE_BASE_PATH}/${baseFolderName}/${homeworkZipName}`;

  await uploadFileToOneDrive(
    env,
    accessToken,
    fileBuffer,
    backupPath,
    'application/zip'
  );
  console.log(`[OneDrive 上传成功] ${homeworkZipName} -> ${backupPath}`);

  // 3. 在内存中解压主 .zip 包
  const { entries } = await unzip(new Blob([fileBuffer]));

  // 4. 遍历主包中的每个学生 .zip 文件
  for (const [studentZipName, studentZipEntry] of Object.entries(entries)) {
    // 确保我们只处理 .zip 文件，并忽略 macOS 的 __MACOSX 文件夹
    if (!studentZipName.endsWith('.zip') || studentZipName.startsWith('__MACOSX/')) {
      console.log(`[跳过] ${studentZipName} (非学生 zip)`);
      continue;
    }

    const studentFolderName = studentZipName.replace(/\.zip$/i, '');
    console.log(`[处理学生] ${studentFolderName}`);

    try {
      // 5. 解压学生 .zip 包
      const studentZipBlob = await studentZipEntry.blob();
      const { entries: studentFiles } = await unzip(studentZipBlob);

      // 6. 准备 Gemini API 的请求内容
      // [ { part: (text | inlineData) }, ... ]
      const geminiParts = [
        { text: getGeminiPrompt(studentFolderName, homeworkZipName) }
      ];

      const studentFileContents = {}; // 用于构建最终报告

      // 7. 遍历学生的所有文件
      for (const [fileName, fileEntry] of Object.entries(studentFiles)) {
        if (fileEntry.isDirectory || fileName.startsWith('__MACOSX/')) continue;

        const fileBlob = await fileEntry.blob();
        const fileBuffer = await fileBlob.arrayBuffer();
        const mimeType = fileBlob.type || 'application/octet-stream';
        const fileContentBase64 = arrayBufferToBase64(fileBuffer); // 转换为 Base64

        // 存储文件内容（用于报告和 Gemini）
        studentFileContents[fileName] = {
          mimeType: mimeType,
          size: fileBlob.size,
        };

        // 检查文件是否能被 Gemini API 接受
        if (isMimeTypeSupportedByGemini(mimeType)) {
          geminiParts.push({ text: `\n--- 学生文件: ${fileName} ---` });
          geminiParts.push({
            inlineData: {
              mimeType: mimeType,
              data: fileContentBase64,
            },
          });
        } else {
          // 如果 Gemini 不支持（例如 .docx, .pdf），我们只告诉它文件名
          geminiParts.push({ text: `\n--- 学生文件: ${fileName} (类型 ${mimeType}, 无法直接读取) ---` });
        }
      }

      // 8. 调用 Gemini API 批改
      console.log(`[Gemini 请求] 正在批改 ${studentFolderName}`);
      const geminiResultText = await callGeminiApi(env, geminiParts);
      console.log(`[Gemini 响应] ${studentFolderName}: ${geminiResultText.substring(0, 50)}...`);

      // 9. 创建最终报告 .json
      const report = {
        student: studentFolderName,
        homework: baseFolderName,
        processedAt: new Date().toISOString(),
        geminiResult: parseGeminiJson(geminiResultText), // 解析 AI 返回的 JSON
        filesSubmitted: studentFileContents,
      };

      // 10. 将批改报告上传回该学生的 OneDrive 文件夹
      const reportPath = `${ONEDRIVE_BASE_PATH}/${baseFolderName}/${studentFolderName}/report.json`;
      await uploadFileToOneDrive(
        env,
        accessToken,
        JSON.stringify(report, null, 2), // 格式化 JSON
        reportPath,
        'application/json'
      );
      console.log(`[报告上传成功] ${reportPath}`);

    } catch (studentErr) {
      console.error(`[处理学生失败] ${studentZipName}: ${studentErr.message}`, studentErr.stack);
      // 可以在此处上传一个 "error_report.json" 到该学生的文件夹
    }
  } // 结束学生循环

  console.log(`[处理完成] ${homeworkZipName}`);
}


// --- 辅助函数 ---

/**
 * [MS Graph] 使用客户端凭据流获取 Access Token
 * @param {object} env - 环境变量
 * @returns {Promise<string|null>} Access Token
 */
async function getMsGraphToken(env) {
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: env.MS_CLIENT_SECRET,
    grant_type: 'client_credentials',
  });

  try {
    const response = await fetch(url, {
      method: 'POST',
      body: body,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      console.error('[MS Token Error]', response.status, errorData);
      throw new Error(`MS Graph Token 请求失败: ${errorData.error_description || response.statusText}`);
    }

    const data = await response.json();
    return data.access_token;
  } catch (err) {
    console.error(`[getMsGraphToken] ${err.message}`);
    return null;
  }
}

/**
 * [OneDrive] 上传文件到 Microsoft Graph
 * @param {object} env - 环境变量
 * @param {string} accessToken - MS Graph Access Token
 * @param {ArrayBuffer|string} fileContent - 文件内容
 * @param {string} pathOnOneDrive - 在 OneDrive 上的完整路径 (例如 'root:/Apps/Homework/file.zip')
 * @param {string} contentType - 文件的 MIME 类型
 * @returns {Promise<object>}
 */
async function uploadFileToOneDrive(env, accessToken, fileContent, pathOnOneDrive, contentType) {
  // 目标用户 ID
  const targetUserID = env.MS_USER_ID;

  // Graph API URL (使用 createUploadSession 适用于大文件，但简单 PUT 适用于 < 4MB)
  // 为简单起见，我们假设文件 < 4MB；对于 > 4MB，需要使用 Upload Session。
  // 路径必须以 ':/' 开头，并以 ':/content' 结尾
  if (!pathOnOneDrive.startsWith('root:')) {
    pathOnOneDrive = `root:${pathOnOneDrive}`;
  }
  const url = `https://graph.microsoft.com/v1.0/users/${targetUserID}/drive/${pathOnOneDrive}:/content`;

  try {
    const response = await fetch(url, {
      method: 'PUT',
      body: fileContent,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': contentType,
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      console.error(`[OneDrive Upload Error] ${pathOnOneDrive}`, response.status, errorData);
      throw new Error(`OneDrive 上传失败 (${pathOnOneDrive}): ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data; // 返回 OneDrive 文件元数据
  } catch (err) {
    console.error(`[uploadFileToOneDrive] ${err.message}`);
    throw err; // 重新抛出错误，让上层捕获
  }
}

/**
 * [Gemini] 调用 Google Gemini API
 * @param {object} env - 环境变量
 * @param {Array<object>} parts - 发送给 Gemini 的内容 (text 和 inlineData)
 * @returns {Promise<string>} Gemini 返回的文本
 */
async function callGeminiApi(env, parts) {
  try {
    const genAI = new GoogleGenerativeAI(env.GOOGLE_API_KEY);
    // 使用 1.5 Flash（如果可用）或最新的模型
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash-latest", safetySettings: GEMINI_SAFETY_SETTINGS });

    const result = await model.generateContent({ contents: [{ role: "user", parts }] });

    // 检查是否有响应
    if (!result.response || !result.response.candidates || result.response.candidates.length === 0) {
      console.error('[Gemini Error] API 未返回有效候选项', result.response);
      throw new Error('Gemini API 未返回有效响应。');
    }
    
    // 检查是否因安全或其他原因被阻止
    const finishReason = result.response.candidates[0].finishReason;
    if (finishReason !== "STOP" && finishReason !== "MAX_TOKENS") {
       console.error(`[Gemini Error] 因 ${finishReason} 停止`, result.response.candidates[0]);
       throw new Error(`Gemini 生成停止: ${finishReason}`);
    }

    // 检查 text() 是否存在
    if (typeof result.response.text !== 'function') {
        // 有时内容在 candidates[0].content.parts[0].text
        if(result.response.candidates[0].content?.parts?.[0]?.text) {
             return result.response.candidates[0].content.parts[0].text;
        }
        console.error('[Gemini Error] 响应格式错误，缺少 text() 方法', result.response);
        throw new Error('Gemini 响应格式错误。');
    }

    return result.response.text();
  } catch (err) {
    console.error(`[callGeminiApi] ${err.message}`, err.stack);
    // 返回一个错误的 JSON，以便上游可以处理
    return JSON.stringify({ error: `调用 Gemini API 失败: ${err.message}` });
  }
}

/**
 * [Gemini] 检查 MIME 类型是否受 Gemini (inlineData) 支持
 * @param {string} mimeType
 * @returns {boolean}
 */
function isMimeTypeSupportedByGemini(mimeType) {
  if (!mimeType) return false;
  // 基于 Google Gemini API 文档
  const supportedImage = ['image/png', 'image/jpeg', 'image/webp', 'image/heic', 'image/heif'];
  const supportedVideo = ['video/mp4', 'video/mpeg', 'video/mov', 'video/avi', 'video/x-flv', 'video/mpg', 'video/webm', 'video/wmv', 'video/3gpp'];
  const supportedAudio = ['audio/wav', 'audio/mp3', 'audio/aiff', 'audio/aac', 'audio/ogg', 'audio/flac'];
  const supportedText = ['text/plain', 'text/html', 'text/css', 'text/javascript', 'application/json', 'text/xml', 'text/csv', 'text/markdown', 'text/rtf'];
  
  // Gemini 1.5 Pro (和 Flash) *不* 支持 PDF, DOCX
  // const supportedDocs = ['application/pdf']; 

  const allSupported = [
      ...supportedImage, 
      ...supportedVideo, 
      ...supportedAudio, 
      ...supportedText
  ];

  return allSupported.includes(mimeType.toLowerCase());
}

/**
 * [Gemini] 构建发送给 Gemini 的主要指令 (Prompt)
 * @param {string} studentName - 学生名
 * @param {string} homeworkName - 作业名
 * @returns {string} Prompt 文本
 */
function getGeminiPrompt(studentName, homeworkName) {
  return `
# 角色
你是一位严格但公平的大学教授，负责批改作业。

# 任务
你将收到来自学生 "${studentName}" 的关于作业 "${homeworkName}" 的文件。
你的任务是：
1.  仔细检查学生提交的所有文件（包括代码、图像、视频、音频和文本）。
2.  根据文件内容和作业要求，给出一个**总分**（0-100）。
3.  提供详细的、有建设性的**总体评语**。
4.  (如果可能) 为每个主要文件或问题点提供**分项反馈**。

# 重要指令
-   **严格的 JSON 输出**：你的回答**必须**是一个严格的 JSON 对象，不能包含任何 markdown 标记 (如 \`\`\`json) 或解释性文本。
-   **基于证据**：你的反馈必须基于学生提交的实际内容。
-   **专业性**：保持专业、学术的语气。

# 输出格式 (必须严格遵守)
{
  "student": "${studentName}",
  "homework": "${homeworkName}",
  "overall_grade": 85,
  "overall_feedback": "学生 ${studentName} 的作业完成度很高。代码逻辑清晰，但缺少对边界情况的处理。请在未来注意代码的稳健性。",
  "detailed_feedback": [
    {
      "file": "main.py",
      "feedback": "主要逻辑正确，但在第 42 行的循环处理存在性能问题。",
      "grade": 90
    },
    {
      "file": "design_document.txt",
      "feedback": "设计文档思路清晰，图表（如果有）表达准确。",
      "grade": 95
    },
    {
      "file": "presentation.mp4",
      "feedback": "视频演示（如果有）流畅，但音频质量有待提高。",
      "grade": 80
    }
  ]
}
`;
}

/**
 * [Util] 解析 Gemini 可能返回的带 \`\`\`json 标记的字符串
 * @param {string} text - Gemini 返回的原始文本
 * @returns {object} 解析后的 JSON 对象
 */
function parseGeminiJson(text) {
  try {
    // 移除 markdown 围栏
    const cleanText = text.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (err) {
    console.warn('[JSON 解析失败] 无法解析 Gemini 响应，将返回原始文本。', text, err);
    // 如果解析失败，返回一个包含原始文本的错误对象
    return {
      error: 'Gemini 返回了非 JSON 格式的响应',
      raw_response: text,
    };
  }
}

/**
 * [Util] 将 ArrayBuffer 转换为 Base64 字符串
 * @param {ArrayBuffer} buffer
 * @returns {string}
 */
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * [HTML] 返回给 GET 请求的上传页面
 * @returns {string} HTML 页面内容
 */
function getHtmlPage() {
  // 我们将 Tailwind CSS 内联到 <style> 标签中，以避免外部 CDN 依赖
  // (在实际生产中，你可能希望从 CDN 加载以利用缓存)
  // 为了简单起见，这里使用 CDN
  return `
<!DOCTYPE html>
<html lang="zh-CN" class="h-full bg-gray-50">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>作业上传</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    body { font-family: 'Inter', sans-serif; }
    .spinner {
      border-color: #f3f3f3;
      border-top-color: #3498db;
      border-radius: 50%;
      width: 24px;
      height: 24px;
      animation: spin 1s linear infinite;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body class="flex items-center justify-center min-h-full p-4">
  <div class="w-full max-w-lg bg-white p-8 rounded-xl shadow-lg">
    <h1 class="text-3xl font-bold text-center text-gray-800 mb-6">作业批改上传系统</h1>
    
    <!-- 上传表单 -->
    <form id="uploadForm" enctype="multipart/form-data">
      <div class="mb-6">
        <label for="homeworkZip" class="block text-sm font-medium text-gray-700 mb-2">选择作业 .zip 包</label>
        <input type="file" name="homeworkZip" id="homeworkZip" accept=".zip" required
               class="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-gray-50 focus:outline-none file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100">
        <p class="mt-1 text-xs text-gray-500">请上传包含所有学生 .zip 文件的**主 .zip 包**。</p>
      </div>

      <button type="submit" id="submitButton" 
              class="w-full flex items-center justify-center px-4 py-3 bg-blue-600 text-white text-base font-medium rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition-colors duration-200">
        <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
          <path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z" clip-rule="evenodd" />
        </svg>
        上传并开始批改
      </button>
    </form>

    <!-- 状态显示区域 -->
    <div id="statusContainer" class="mt-6 text-center">
      <!-- 加载指示器 (默认隐藏) -->
      <div id="loadingIndicator" class="hidden flex-col items-center justify-center">
        <div class="spinner"></div>
        <p class="text-sm text-gray-600 mt-2">正在上传文件，请勿关闭页面...</p>
      </div>
      <!-- 消息区域 -->
      <div id="messageArea" class="text-sm p-4 rounded-lg hidden"></div>
    </div>

  </div>

  <script>
    const form = document.getElementById('uploadForm');
    const submitButton = document.getElementById('submitButton');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const messageArea = document.getElementById('messageArea');

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const fileInput = document.getElementById('homeworkZip');
      if (!fileInput.files || fileInput.files.length === 0) {
        showMessage('请先选择一个文件。', 'error');
        return;
      }

      // 禁用按钮并显示加载
      submitButton.disabled = true;
      submitButton.classList.add('bg-gray-400', 'cursor-not-allowed');
      loadingIndicator.classList.remove('hidden');
      loadingIndicator.classList.add('flex');
      messageArea.classList.add('hidden');

      const formData = new FormData(form);

      try {
        const response = await fetch(window.location.href, { // POST 到当前 URL
          method: 'POST',
          body: formData,
        });

        const responseText = await response.text();

        if (response.ok) {
          // 状态 200-299 (特别是 202 Accepted)
          showMessage(responseText, 'success');
          form.reset(); // 清空表单
        } else {
          // 状态 400-599
          showMessage(responseText, 'error');
        }

      } catch (error) {
        console.error('上传失败:', error);
        showMessage('上传失败，请检查网络连接或联系管理员。 ' + error.message, 'error');
      } finally {
        // 恢复按钮和隐藏加载
        submitButton.disabled = false;
        submitButton.classList.remove('bg-gray-400', 'cursor-not-allowed');
        loadingIndicator.classList.add('hidden');
        loadingIndicator.classList.remove('flex');
      }
    });

    function showMessage(message, type) {
      messageArea.textContent = message;
      messageArea.classList.remove('hidden', 'bg-green-100', 'text-green-800', 'bg-red-100', 'text-red-800');
      
      if (type === 'success') {
        messageArea.classList.add('bg-green-100', 'text-green-800');
      } else if (type === 'error') {
        messageArea.classList.add('bg-red-100', 'text-red-800');
      }
      messageArea.classList.remove('hidden');
    }
  </script>
</body>
</html>
  `;
}

