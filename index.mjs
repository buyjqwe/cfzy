/**
 * Cloudflare Worker: 自动批改作业
 * ---------------------------------
 * 认证流程: 客户端凭据流 (Client Credentials Flow)
 *
 * 必需的 Secrets:
 * - GOOGLE_API_KEY: Gemini API 密钥
 * - MS_CLIENT_ID: Microsoft Azure App Client ID
 * - MS_CLIENT_SECRET: Microsoft Azure App Client Secret
 * - MS_TENANT_ID: Microsoft Azure Tenant ID (租户ID)
 * - MS_USER_ID: 要上传到的目标 OneDrive 账户的用户 ID (或 User Principal Name, 如 'user@domain.com')
 */

import { unzip } from 'unzipit'; // 用于解压 .zip 文件
import { GoogleGenerativeAI } from '@google/generative-ai'; // Gemini API

// --- OneDrive (Microsoft Graph) 核心功能 ---

/**
 * @description 使用 Client Credentials Flow 获取 Access Token
 * 这是服务器到服务器的身份验证，代表“应用本身”
 * @param {object} env - Worker 的环境变量 (Secrets)
 * @returns {Promise<string>} - Access Token
 */
async function getAccessToken(env) {
  const { MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID } = env;

  if (!MS_CLIENT_ID || !MS_CLIENT_SECRET || !MS_TENANT_ID) {
    throw new Error('Microsoft Graph API 凭据 (ID, Secret, Tenant) 未在 Secrets 中配置。');
  }

  const tokenEndpoint = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
  
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: MS_CLIENT_ID,
      client_secret: MS_CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default', // 请求应用本身被授予的所有权限
    }),
  });

  if (!response.ok) {
    const errorData = await response.json();
    console.error('获取 Access Token 失败:', errorData);
    throw new Error(`获取 Access Token 失败: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  return data.access_token;
}

/**
 * @description 上传文件到指定的 OneDrive 路径
 * @param {ArrayBuffer} fileContent - 文件的二进制内容
 * @param {string} targetPath - 目标 OneDrive 路径 (例如: "Homework/student1.zip")
 * @param {string} accessToken - Microsoft Graph API Access Token
 * @param {string} userId - 目标 OneDrive 账户的用户 ID 或 UPN
 * @returns {Promise<Response>}
 */
async function uploadToOneDrive(fileContent, targetPath, accessToken, userId) {
  // 注意: 即使使用 Client Credentials，我们仍然需要指定要操作 *哪个用户* 的 Drive
  const endpoint = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${targetPath}:/content`;
  
  const response = await fetch(endpoint, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/zip', // 假设我们总是上传 zip
    },
    body: fileContent,
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`OneDrive 上传失败 (${targetPath}):`, errorText);
    throw new Error(`OneDrive 上传失败: ${response.status} ${response.statusText}`);
  }

  return response.json(); // 返回上传成功的元数据
}

// --- Gemini (AI 批改) 核心功能 ---

/**
 * @description 初始化 Gemini 客户端
 * @param {object} env - Worker 的环境变量 (Secrets)
 * @returns {GenerativeModel}
 */
function getGeminiModel(env) {
  if (!env.GOOGLE_API_KEY) {
    throw new Error('GOOGLE_API_KEY 未在 Secrets 中配置。');
  }
  const genAI = new GoogleGenerativeAI(env.GOOGLE_API_KEY);
  return genAI.getGenerativeModel({ model: 'gemini-2.5-flash-preview-09-2025' }); // 使用 Flash 模型
}

/**
 * @description （占位符）调用 Gemini API 批改作业
 * @param {object} studentFiles - 包含学生文件内容的对象 (例如: { 'main.py': '...' })
 * @param {object} model - Gemini 模型实例
 * @returns {Promise<object>} - 批改结果
 */
async function gradeHomeworkWithGemini(studentFiles, model) {
  // TODO: 构建一个更复杂的 prompt，包含所有文件内容
  
  // 简化示例：只发送第一个文件的内容
  const firstFileName = Object.keys(studentFiles)[0];
  const firstFileContent = studentFiles[firstFileName];

  if (!firstFileContent) {
    return { error: '未找到可批改的文件。' };
  }

  const prompt = `
    # 角色：你是一个 Python 课程的助教。
    # 任务：批改以下学生的作业文件。
    
    ## 文件名：${firstFileName}
    
    ## 文件内容：
    \`\`\`python
    ${firstFileContent}
    \`\`\`
    
    # 批改要求：
    1. 总结代码的功能。
    2. 指出任何明显的错误或可以改进的地方。
    3. 给出一个 0-100 的分数。
    4. 以 JSON 格式返回你的批改结果。
    
    {"score": ..., "feedback": "..."}
  `;

  try {
    const result = await model.generateContent(prompt);
    const response = await result.response;
    const text = response.text();

    // 粗略的 JSON 解析
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    } else {
      return { error: 'AI 未返回有效的 JSON 格式。', raw: text };
    }
  } catch (error) {
    console.error('Gemini API 调用失败:', error);
    return { error: 'Gemini API 调用失败。' };
  }
}

// --- Worker 入口 ---

export default {
  /**
   * @param {Request} request - 传入的 HTTP 请求
   * @param {object} env - Worker 的环境变量 (Secrets)
   * @param {object} ctx - Worker 的执行上下文
   * @returns {Promise<Response>}
   */
  async fetch(request, env, ctx) {
    if (request.method !== 'POST') {
      return new Response('无效的请求方法。请使用 POST 上传文件。', { status: 405 });
    }

    if (!env.MS_USER_ID) {
      return new Response('MS_USER_ID 未在 Secrets 中配置。', { status: 500 });
    }

    try {
      // 1. 获取 Microsoft Graph API 的 Access Token
      const accessToken = await getAccessToken(env);

      // 2. 从请求中获取主作业包 (homework.zip)
      const mainZipBuffer = await request.arrayBuffer();
      if (!mainZipBuffer || mainZipBuffer.byteLength === 0) {
        return new Response('未找到上传的文件。', { status: 400 });
      }

      // 3. (异步) 上传主包到 OneDrive 备份
      const timestamp = new Date().toISOString().replace(/:/g, '-');
      const mainZipPath = `Homework_Submissions/Backup_${timestamp}.zip`;
      
      // 我们不等待这个完成，让它在后台运行
      ctx.waitUntil(uploadToOneDrive(mainZipBuffer, mainZipPath, accessToken, env.MS_USER_ID));
      
      // 4. 初始化 Gemini
      const geminiModel = getGeminiModel(env);

      // 5. 解压主包 (homework.zip)
      const { entries: studentZips } = await unzip(new Uint8Array(mainZipBuffer));
      
      const gradingResults = {};
      const processingPromises = [];

      // 6. 遍历主包中的每个学生 .zip 文件
      for (const [zipName, zipEntry] of Object.entries(studentZips)) {
        if (zipEntry.isDirectory || !zipName.endsWith('.zip')) {
          continue; // 跳过文件夹或非 zip 文件
        }

        // 启动一个并行的处理任务
        const processTask = async () => {
          try {
            // 7. 解压学生 .zip 文件 (例如 student_A.zip)
            const studentZipBuffer = await zipEntry.arrayBuffer();
            const { entries: studentFiles } = await unzip(new Uint8Array(studentZipBuffer));
            
            const studentFileContents = {};
            
            // 8. 读取学生的所有作业文件内容
            for (const [fileName, fileEntry] of Object.entries(studentFiles)) {
              if (fileEntry.isDirectory) continue;
              // 我们只处理文本类文件进行批改
              if (/\.(py|txt|md|js|html|css|java|c|cpp)$/i.test(fileName)) {
                studentFileContents[fileName] = await fileEntry.text();
              }
            }

            // 9. (异步) 调用 Gemini API 批改
            const grade = await gradeHomeworkWithGemini(studentFileContents, geminiModel);
            gradingResults[zipName] = grade;

          } catch (unzipError) {
            console.error(`处理 ${zipName} 失败:`, unzipError);
            gradingResults[zipName] = { error: `解压学生 ${zipName} 失败。` };
          }
        };
        processingPromises.push(processTask());
      }

      // 10. 等待所有批改任务完成
      await Promise.all(processingPromises);

      // 11. 返回所有学生的批改结果
      return new Response(JSON.stringify(gradingResults, null, 2), {
        headers: { 'Content-Type': 'application/json' },
      });

    } catch (error) {
      console.error('Worker 执行时发生严重错误:', error);
      return new Response(error.message, { status: 500 });
    }
  },
};

