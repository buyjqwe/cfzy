/**
 * Cloudflare Worker for Secure Homework Grading
 *
 * Features:
 * 1. Email-based login with verification code (using MS Graph API for mail, KV for code storage).
 * 2. Session management via secure, HttpOnly cookies.
 * 3. Protected routes:
 * - GET /: Shows login page or main app (upload/download).
 * - POST /: Handles homework zip upload (protected).
 * - GET /summary: Handles summary CSV download (protected).
 * - POST /login: Creates and emails a verification code.
 * - POST /verify: Verifies code and sets session cookie.
 * - POST /logout: Clears session cookie.
 * 4. Background processing (ctx.waitUntil) for:
 * - Uploading main zip to OneDrive.
 * - Unzipping, processing student files.
 * - Calling Gemini API for grading.
 * - Uploading individual JSON reports to OneDrive.
 * 5. On-demand CSV summary generation by iterating OneDrive folders.
 */

import { unzip } from 'unzipit'; // 用于解压 .zip 文件
import { GoogleGenerativeAI } from '@google/generative-ai'; // Gemini API

// --- 常量 ---
const SESSION_COOKIE_NAME = '__auth_session';
const SESSION_DURATION_SECONDS = 7 * 24 * 60 * 60; // 7 天
const VERIFICATION_CODE_TTL_SECONDS = 300; // 5 分钟

/**
 * 主入口点: 处理所有传入的 HTTP 请求
 */
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const path = url.pathname;
    const method = request.method;

    // 1. 解析会话 Cookie
    const session = await parseSessionCookie(request, env);

    try {
      // --- 公共路由 (无需登录) ---
      if (method === 'POST' && path === '/login') {
        return await handleLoginRequest(request, env);
      }
      if (method === 'POST' && path === '/verify') {
        return await handleVerifyRequest(request, env);
      }
      if (method === 'POST' && path === '/logout') {
        return handleLogoutRequest();
      }

      // --- 路由守卫: 检查是否登录 ---
      if (!session.valid) {
        // 如果未登录，无论是访问主页还是受保护的API，都显示登录页面
        if (method === 'GET' && path === '/') {
          return new Response(getHtmlPage('login', null, env.APP_TITLE || '作业批改系统'), {
            headers: { 'Content-Type': 'text/html; charset=utf-8' },
          });
        }
        // 对 API 请求返回 401
        return new Response('Unauthorized', { status: 401 });
      }

      // --- 受保护的路由 (需要登录) ---
      if (method === 'GET' && path === '/') {
        // 已登录，显示主应用页面 (上传/下载)
        return new Response(getHtmlPage('app', session.email, env.APP_TITLE || '作业批改系统'), {
          headers: { 'Content-Type': 'text/html; charset=utf-8' },
        });
      }

      if (method === 'POST' && path === '/') {
        // 处理作业上传
        return await handleUpload(request, env, ctx);
      }

      if (method === 'GET' && path === '/summary') {
        // 处理成绩汇总下载
        const homeworkName = url.searchParams.get('homework');
        if (!homeworkName) {
          return new Response('缺少 "homework" 参数', { status: 400 });
        }
        return await handleSummaryDownload(env, homeworkName);
      }

      // 404 Not Found
      return new Response('Not Found', { status: 404 });
    } catch (err) {
      console.error('Fetch Handler Error:', err.stack);
      return new Response(`服务器内部错误: ${err.message}`, { status: 500 });
    }
  },
};

// =================================================================
// --- 认证与会话 (Auth & Session) ---
// =================================================================

/**
 * 1. 处理登录请求 (/login)
 * 生成验证码, 存入 KV, 并发送邮件
 */
async function handleLoginRequest(request, env) {
  const formData = await request.formData();
  const email = formData.get('email');

  if (!email || !email.includes('@')) { // 简单的邮箱验证
    return new Response('请输入有效的邮箱地址', { status: 400 });
  }

  // 检查是否在白名单中
  const WHITELIST = (env.TEACHER_WHITELIST || '').split(',').map(e => e.trim()).filter(Boolean);
  if (WHITELIST.length > 0 && !WHITELIST.includes(email)) {
    console.warn(`[Login Attempt] Blocked non-whitelist email: ${email}`);
    // 故意返回通用成功消息，防止探测白名单
    return new Response('如果邮箱地址在白名单中，验证码已发送。', { status: 200 });
  }
  
  const code = Math.floor(100000 + Math.random() * 900000).toString(); // 6位验证码
  const key = `authcode:${email}`;

  await env.AUTH_KV.put(key, code, { expirationTtl: VERIFICATION_CODE_TTL_SECONDS });

  const emailSubject = `您的登录验证码: ${code}`;
  const emailBody = `
    <html>
      <body>
        <h3>欢迎使用 ${env.APP_TITLE || '作业批改系统'}</h3>
        <p>您的登录验证码是： <h1>${code}</h1></p>
        <p>此验证码将在 ${VERIFICATION_CODE_TTL_SECONDS / 60} 分钟内有效。</p>
        <p>如果您未请求此验证码，请忽略此邮件。</p>
      </body>
    </html>
  `;

  const token = await getMsGraphToken(env);
  if (!token) {
    return new Response('无法获取认证令牌以发送邮件', { status: 500 });
  }

  const sent = await sendEmail(env, token, email, emailSubject, emailBody);
  if (!sent) {
    return new Response('发送邮件失败', { status: 500 });
  }

  return new Response('验证码已发送至您的邮箱，请查收。', { status: 200 });
}

/**
 * 2. 处理验证码验证请求 (/verify)
 * 验证 KV 中的代码, 成功则设置会话 Cookie
 */
async function handleVerifyRequest(request, env) {
  const formData = await request.formData();
  const email = formData.get('email');
  const code = formData.get('code');

  if (!email || !code) {
    return new Response('邮箱或验证码不能为空', { status: 400 });
  }

  const key = `authcode:${email}`;
  const storedCode = await env.AUTH_KV.get(key);

  if (!storedCode) {
    return new Response('验证码已过期或不存在', { status: 400 });
  }

  if (storedCode !== code) {
    return new Response('验证码错误', { status: 400 });
  }

  // 验证成功, 删除验证码
  await env.AUTH_KV.delete(key);

  // 创建会话
  const sessionId = crypto.randomUUID();
  const sessionKey = `session:${sessionId}`;
  const sessionData = JSON.stringify({ email: email, createdAt: Date.now() });

  await env.AUTH_KV.put(sessionKey, sessionData, { expirationTtl: SESSION_DURATION_SECONDS });

  // 设置 HttpOnly, Secure Cookie
  const cookie = `${SESSION_COOKIE_NAME}=${sessionId}; HttpOnly; Secure; Path=/; Max-Age=${SESSION_DURATION_SECONDS}; SameSite=Strict`;

  // 重定向到主页
  return new Response(null, {
    status: 302,
    headers: {
      'Set-Cookie': cookie,
      'Location': '/',
    },
  });
}

/**
 * 3. 处理登出请求 (/logout)
 * 清除会话 Cookie 和 KV 中的会话数据
 */
function handleLogoutRequest() {
  // 发送一个立即过期的 Cookie 来清除它
  const cookie = `${SESSION_COOKIE_NAME}=""; HttpOnly; Secure; Path=/; Max-Age=0; SameSite=Strict`;
  // (注意: 我们没有删除 KV 中的会话, 它会自动过期。如果需要立即失效, 还需要从请求中解析 sessionId 并删除)
  return new Response(null, {
    status: 302,
    headers: {
      'Set-Cookie': cookie,
      'Location': '/',
    },
  });
}

/**
 * 4. 解析会话 Cookie
 * 检查 Cookie 有效性并从 KV 返回会话数据
 */
async function parseSessionCookie(request, env) {
  const cookieHeader = request.headers.get('Cookie');
  if (!cookieHeader) {
    return { valid: false };
  }

  const cookies = cookieHeader.split(';').map(c => c.trim());
  const sessionCookie = cookies.find(c => c.startsWith(`${SESSION_COOKIE_NAME}=`));

  if (!sessionCookie) {
    return { valid: false };
  }

  const sessionId = sessionCookie.split('=')[1];
  if (!sessionId) {
    return { valid: false };
  }

  const sessionKey = `session:${sessionId}`;
  const sessionData = await env.AUTH_KV.get(sessionKey);

  if (!sessionData) {
    return { valid: false }; // 会话过期或无效
  }

  try {
    const data = JSON.parse(sessionData);
    return { valid: true, email: data.email, sessionId: sessionId };
  } catch (e) {
    return { valid: false };
  }
}

// =================================================================
// --- 核心应用逻辑 (App Logic) ---
// =================================================================

/**
 * 5. 处理文件上传 (POST /)
 * 立即返回响应, 并在后台启动处理
 */
async function handleUpload(request, env, ctx) {
  const formData = await request.formData();
  const homeworkFile = formData.get('homeworkFile');

  // [BUG FIX] 检查文件是否存在
  if (!homeworkFile || !homeworkFile.name || homeworkFile.name === 'undefined') {
    return new Response('未选择任何文件。请选择一个 .zip 文件后重试。', { status: 400 });
  }
  
  // 移除 .zip 后缀作为作业名称
  const homeworkName = homeworkFile.name.replace(/\.zip$/i, '');
  const fileBuffer = await homeworkFile.arrayBuffer();

  // 关键: 立即返回 202 Accepted 响应, 防止浏览器超时
  ctx.waitUntil(
    processZipInBackground(env, fileBuffer, homeworkName)
  );

  return new Response(`文件 "${homeworkFile.name}" 已收到，正在后台处理批改。这可能需要几分钟时间。`, { status: 202 });
}

/**
 * 6. 处理成绩汇总下载 (GET /summary)
 * 按需生成并返回 CSV 文件
 */
async function handleSummaryDownload(env, homeworkName) {
  console.log(`[Summary] 开始为 "${homeworkName}" 生成汇总...`);
  const token = await getMsGraphToken(env);
  if (!token) return new Response('无法获取 MS Token', { status: 500 });

  const homeworkFolderPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}`;
  const studentFolders = await listStudentFolders(env, token, homeworkFolderPath);

  if (studentFolders.length === 0) {
    return new Response(`未找到作业 "${homeworkName}" 的任何已处理学生数据。请确认名称是否正确, 以及AI是否已完成批改。`, { status: 404 });
  }

  const summaryData = [];
  
  // 并行抓取所有 report.json
  const reportPromises = studentFolders.map(async (folder) => {
    try {
      const reportPath = `${homeworkFolderPath}/${folder.name}/report.json`;
      const reportJson = await getOneDriveFileContent(env, token, reportPath);
      
      if (reportJson) {
        return {
          studentName: folder.name, // 文件夹名即学生名
          grade: reportJson.overall_grade,
          feedback: reportJson.overall_feedback,
        };
      } else {
        return { studentName: folder.name, grade: "N/A", feedback: "report.json 未找到" };
      }
    } catch (e) {
      console.error(`[Summary] 处理 ${folder.name} 失败:`, e.message);
      return { studentName: folder.name, grade: "N/A", feedback: `处理失败: ${e.message}` };
    }
  });

  const reports = await Promise.all(reportPromises);

  // 添加尚未完成批改的学生 (如果需要)
  // (当前逻辑: 仅显示已生成报告的学生)
  
  // 过滤掉批改未完成的 (如果文件夹存在但 report.json 不存在)
  const finalReports = reports.filter(r => r.grade !== "N/A");
  
  if (finalReports.length === 0) {
      return new Response(`作业 "${homeworkName}" 已找到, 但AI尚未完成任何批改。`, { status: 404 });
  }

  // 生成 CSV
  const csvContent = generateCsv(finalReports);
  const safeFileName = homeworkName.replace(/[^a-z0-9]/gi, '_');

  return new Response(csvContent, {
    headers: {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': `attachment; filename="${safeFileName}_summary.csv"`,
    },
  });
}

// =================================================================
// --- 后台处理 (Background Processing) ---
// =================================================================

/**
 * 7. (后台) 主处理流程
 */
async function processZipInBackground(env, fileBuffer, homeworkName) {
  console.log(`[后台处理] 开始: ${homeworkName}`);

  try {
    // 1. 获取 MS Token
    const token = await getMsGraphToken(env);
    if (!token) {
      console.error('[后台处理] 失败: 无法获取 MS Token');
      return;
    }
    console.log('[后台处理] MS Token 成功获取');

    // 2. (可选) 上传原始 Zip 包作为备份
    const zipPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}/${homeworkName}_archive.zip`;
    await uploadToOneDrive(env, token, zipPath, fileBuffer, 'application/zip');
    console.log(`[后台处理] 原始 ZIP 备份成功: ${zipPath}`);

    // 3. 从 Zip 中解析学生文件
    const students = await getStudentsFromZip(fileBuffer);
    if (students.length === 0) {
      console.warn('[后台处理] 警告: Zip 包中未找到学生文件。');
      return;
    }
    console.log(`[后台处理] 解析到 ${students.length} 个学生`);

    // 4. 循环处理每个学生 (并行)
    const geminiApiKey = env.GOOGLE_API_KEY;
    const model = new GoogleGenerativeAI(geminiApiKey).getGenerativeModel({ model: 'gemini-1.5-flash' });

    const processingTasks = students.map(student => 
      processSingleStudent(env, token, model, homeworkName, student)
    );
    
    await Promise.allSettled(processingTasks);
    
    console.log(`[后台处理] 完成: ${homeworkName}`);

  } catch (err) {
    console.error(`[后台处理] 致命错误 (${homeworkName}):`, err.stack);
  }
}

/**
 * 8. (后台) 处理单个学生
 */
async function processSingleStudent(env, token, model, homeworkName, student) {
  const studentName = student.name;
  console.log(`[处理学生] 开始: ${studentName}`);

  try {
    // 1. 准备 Gemini 请求
    const prompt = env.GEMINI_PROMPT || '请批改这份作业。';
    const requestParts = [prompt, `学生: ${studentName}`];
    
    for (const file of student.files) {
      requestParts.push(`--- 文件: ${file.name} ---`);
      requestParts.push({
        inlineData: {
          data: btoa(String.fromCharCode.apply(null, file.data)),
          mimeType: file.mimeType,
        },
      });
    }

    // 2. 调用 Gemini API
    console.log(`[Gemini 请求] 正在批改: ${studentName}`);
    const result = await model.generateContent({ contents: [{ role: 'user', parts: requestParts }] });
    const responseText = result.response.text();
    console.log(`[Gemini 响应] ${studentName} 批改完成`);

    // 3. 解析 Gemini 响应 (假设为 JSON)
    let reportJsonText = responseText;
    if (responseText.startsWith('```json')) {
      reportJsonText = responseText.substring(7, responseText.length - 3).trim();
    }
    
    // 验证 JSON 格式
    let reportData;
    try {
        reportData = JSON.parse(reportJsonText);
    } catch (e) {
        console.warn(`[Gemini 警告] ${studentName} 的响应不是有效JSON, 将作为纯文本保存。内容: ${responseText.substring(0, 50)}...`);
        // 创建一个回退的 JSON
        reportData = {
            overall_grade: "N/A (AI未返回JSON)",
            overall_feedback: responseText,
            detailed_grades: []
        };
        reportJsonText = JSON.stringify(reportData, null, 2);
    }

    // 4. 上传批改报告到 OneDrive
    const reportPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}/${studentName}/report.json`;
    await uploadToOneDrive(env, token, reportPath, reportJsonText, 'application/json');
    console.log(`[报告上传成功] ${reportPath}`);

  } catch (err) {
    console.error(`[处理学生] ${studentName} 失败:`, err.stack);
    // (可选) 上传一个错误报告
    const errorReport = { error: err.message, stack: err.stack };
    const errorPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}/${studentName}/error.json`;
    await uploadToOneDrive(env, token, errorPath, JSON.stringify(errorReport, null, 2), 'application/json');
  }
}

// =================================================================
// --- 辅助工具 (Utilities) ---
// =================================================================

/**
 * 9. (工具) 从主 Zip 中解析学生文件
 * 假设: 顶层 zip 包含每个学生的 zip (e.g., "张三.zip", "李四.zip")
 */
async function getStudentsFromZip(zipBuffer) {
  const { entries } = await unzip(zipBuffer);
  const students = [];
  const studentZipEntries = Object.values(entries).filter(entry => entry.name.endsWith('.zip') && !entry.name.startsWith('__MACOSX'));

  for (const entry of studentZipEntries) {
    const studentName = entry.name.replace(/\.zip$/i, '').split('/').pop(); // "folder/张三.zip" -> "张三"
    const studentZipBuffer = await entry.arrayBuffer();
    const { entries: studentFiles } = await unzip(studentZipBuffer);
    
    const filesData = [];
    for (const fileEntry of Object.values(studentFiles)) {
      if (fileEntry.isDirectory || fileEntry.name.startsWith('__MACOSX') || fileEntry.name.endsWith('.DS_Store')) {
        continue;
      }
      const fileData = await fileEntry.arrayBuffer();
      filesData.push({
        name: fileEntry.name,
        data: new Uint8Array(fileData),
        mimeType: fileEntry.contentType || 'application/octet-stream',
      });
    }

    if (filesData.length > 0) {
      students.push({ name: studentName, files: filesData });
    }
  }
  return students;
}

/**
 * 10. (工具) 生成 CSV 内容
 */
function generateCsv(data) {
  if (data.length === 0) return '';
  
  const headers = ['StudentName', 'Grade', 'Feedback'];
  const csvRows = [headers.join(',')]; // Header row

  for (const row of data) {
    const studentName = `"${(row.studentName || '').replace(/"/g, '""')}"`;
    const grade = row.grade;
    const feedback = `"${(row.feedback || '').replace(/"/g, '""').replace(/\n/g, ' ')}"`; // 替换换行符
    csvRows.push([studentName, grade, feedback].join(','));
  }
  
  // 添加 UTF-8 BOM 头, 确保 Excel 正确打开中文
  return '\uFEFF' + csvRows.join('\n');
}

// =================================================================
// --- Microsoft Graph API 助手 ---
// =================================================================

/**
 * 11. (MS API) 获取 Access Token (Client Credentials Flow)
 */
async function getMsGraphToken(env) {
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET,
    scope: '[https://graph.microsoft.com/.default](https://graph.microsoft.com/.default)',
    grant_type: 'client_credentials',
  });

  const response = await fetch(url, {
    method: 'POST',
    body: body,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[MS Token 失败]', response.status, errorText);
    return null;
  }

  const data = await response.json();
  return data.access_token;
}

/**
 * 12. (MS API) 发送邮件
 * 需要 Mail.Send (Application) 权限
 */
async function sendEmail(env, token, toEmail, subject, htmlBody) {
  // 注意: "saveToSentItems" 必须为 "false"
  // 否则需要 Mail.ReadWrite (Application) 权限
  const emailData = {
    message: {
      subject: subject,
      body: {
        contentType: 'HTML',
        content: htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: toEmail,
          },
        },
      ],
    },
    saveToSentItems: 'false',
  };

  const response = await fetch(`https://graph.microsoft.com/v1.0/users/${env.MS_USER_ID}/sendMail`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(emailData),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[Email 失败]', response.status, errorText);
    return false;
  }
  return true;
}

/**
 * 13. (MS API) 上传文件到 OneDrive
 * (自动创建父文件夹, 覆盖同名文件)
 */
async function uploadToOneDrive(env, token, pathOnOneDrive, content, contentType) {
  // Graph API 使用 /items/root:/path/to/file:/content
  const url = `https://graph.microsoft.com/v1.0/users/${env.MS_USER_ID}/drive/root:${pathOnOneDrive}:/content`;
  
  const response = await fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': contentType,
    },
    body: content,
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`[OneDrive 上传失败] ${pathOnOneDrive}:`, response.status, errorText);
  }
  return response.ok;
}

/**
 * 14. (MS API) 列出文件夹中的子项 (用于 /summary)
 */
async function listStudentFolders(env, token, folderPath) {
  const url = `https://graph.microsoft.com/v1.0/users/${env.MS_USER_ID}/drive/root:${folderPath}:/children?$select=name,folder`;
  
  try {
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` },
    });
    if (!response.ok) {
      if (response.status === 404) {
        console.warn(`[listFolders] 404 - 文件夹未找到: ${folderPath}`);
        return [];
      }
      throw new Error(`Graph API error ${response.status}: ${await response.text()}`);
    }
    const data = await response.json();
    // 仅返回文件夹
    return data.value.filter(item => item.folder);
  } catch (e) {
    console.error(`[listFolders] 失败: ${folderPath}`, e.message);
    return [];
  }
}

/**
 * 15. (MS API) 获取 OneDrive 上的文件内容 (用于 /summary)
 */
async function getOneDriveFileContent(env, token, filePath) {
  const url = `https://graph.microsoft.com/v1.0/users/${env.MS_USER_ID}/drive/root:${filePath}:/content`;
  
  try {
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` },
    });
    if (!response.ok) {
      if (response.status === 404) return null; // 文件不存在
      throw new Error(`Graph API error ${response.status}: ${await response.text()}`);
    }
    // 假设 report.json 总是 JSON
    return await response.json(); 
  } catch (e) {
    console.error(`[getFile] 失败: ${filePath}`, e.message);
    return null;
  }
}


// =================================================================
// --- HTML 页面模板 ---
// =================================================================

/**
 * 16. (工具) 生成动态 HTML 页面
 */
function getHtmlPage(pageType, email = null, appTitle = "作业批改系统") {
  const title = `<h1>${appTitle}</h1>`;
  
  const commonStyles = `
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 2rem; background-color: #f4f7f6; display: flex; justify-content: center; align-items: center; min-height: 90vh; }
      main { background-color: #ffffff; border-radius: 12px; box-shadow: 0 8px 20px rgba(0,0,0,0.08); padding: 2.5rem; max-width: 600px; width: 100%; box-sizing: border-box; }
      h1, h2, h3 { color: #333; margin-top: 0; }
      h1 { text-align: center; color: #1a1a1a; margin-bottom: 2rem; }
      h2 { border-bottom: 2px solid #eee; padding-bottom: 10px; margin-top: 2rem; }
      form { display: flex; flex-direction: column; gap: 1rem; }
      input[type="email"], input[type="text"], input[type="file"] { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; font-size: 1rem; }
      button { background-color: #007aff; color: white; border: none; padding: 14px 20px; border-radius: 8px; font-size: 1rem; font-weight: 600; cursor: pointer; transition: background-color 0.2s; }
      button:hover { background-color: #0056b3; }
      button[type="submit"][name="logout"] { background-color: #f44336; margin-top: 1rem; }
      button[type="submit"][name="logout"]:hover { background-color: #d32f2f; }
      .user-info { display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; background: #eef; padding: 10px 15px; border-radius: 8px; font-size: 0.9rem; }
      .message { padding: 1rem; border-radius: 8px; margin-top: 1rem; text-align: center; font-weight: 500; }
      .message.success { background-color: #d4edda; color: #155724; }
      .message.error { background-color: #f8d7da; color: #721c24; }
      .message.loading { background-color: #fff3cd; color: #856404; }
    </style>
  `;

  let bodyContent = '';

  if (pageType === 'login') {
    bodyContent = `
      ${title}
      <div id="login-box">
        <h2>教师登录</h2>
        <form id="login-form">
          <label for="email">教师邮箱:</label>
          <input type="email" id="email" name="email" required placeholder="name@example.com">
          <button type="submit">发送验证码</button>
        </form>
      </div>
      <div id="verify-box" style="display:none;">
        <h2>验证码</h2>
        <p>验证码已发送至 <strong id="email-display"></strong></p>
        <form id="verify-form">
          <input type="hidden" id="email-hidden" name="email">
          <label for="code">6位验证码:</label>
          <input type="text" id="code" name="code" required pattern="\\d{6}" maxlength="6">
          <button type="submit">登录</button>
        </form>
      </div>
      <div id="message-container"></div>
    `;
  } else if (pageType === 'app') {
    bodyContent = `
      <div class="user-info">
        <span>已登录: <strong>${email || '教师'}</strong></span>
        <form action="/logout" method="POST" style="margin:0;">
          <button type="submit" name="logout" style="padding: 8px 12px; font-size: 0.9rem;">退出登录</button>
        </form>
      </div>
      ${title}
      
      <!-- 1. 上传作业 -->
      <section id="upload-section">
        <h2>1. 上传作业压缩包</h2>
        <form id="upload-form" enctype="multipart/form-data">
          <label for="homeworkFile">选择作业 .zip 包:</label>
          <p style="font-size:0.85rem; color:#555;">(压缩包内应包含每个学生的 .zip 文件)</p>
          <input type="file" id="homeworkFile" name="homeworkFile" accept=".zip" required>
          <button type="submit">上传并开始批改</button>
        </form>
      </section>

      <!-- 2. 下载总表 -->
      <section id="download-section">
        <h2>2. 下载成绩总表</h2>
        <form id="download-form">
          <label for="homeworkName">输入作业名称:</label>
          <p style="font-size:0.85rem; color:#555;">(必须与您上传的 .zip 文件名(不含.zip)完全一致)</p>
          <input type="text" id="homeworkName" name="homeworkName" required placeholder="例如: 22计网1-GET+HEAD">
          <button type="submit">下载 CSV 汇总表</button>
        </form>
      </section>

      <div id="message-container"></div>
    `;
  }

  const scripts = `
    <script>
      function showMessage(type, text) {
        const container = document.getElementById('message-container');
        if (!container) return;
        container.innerHTML = \`<div class="message \${type}">\${text}</div>\`;
      }

      // --- 登录/验证逻辑 ---
      const loginForm = document.getElementById('login-form');
      const verifyForm = document.getElementById('verify-form');
      if (loginForm && verifyForm) {
        loginForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          showMessage('loading', '正在发送验证码...');
          const formData = new FormData(loginForm);
          const email = formData.get('email');
          
          const response = await fetch('/login', { method: 'POST', body: formData });
          const text = await response.text();
          
          if (response.ok) {
            showMessage('success', text);
            document.getElementById('login-box').style.display = 'none';
            document.getElementById('verify-box').style.display = 'block';
            document.getElementById('email-display').textContent = email;
            document.getElementById('email-hidden').value = email;
          } else {
            showMessage('error', text);
          }
        });

        // 验证表单使用标准 POST 提交, 因为成功后会 302 重定向并设置 Cookie
        // (不需要 JS 拦截)
      }

      // --- 上传逻辑 ---
      const uploadForm = document.getElementById('upload-form');
      if (uploadForm) {
        uploadForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          showMessage('loading', '正在上传文件，请稍候...');
          const formData = new FormData(uploadForm);
          
          try {
            const response = await fetch('/', { method: 'POST', body: formData });
            const text = await response.text();
            
            if (response.ok) { // 200-299 (包括 202)
              showMessage('success', text);
              uploadForm.reset();
            } else {
              showMessage('error', \`上传失败: \${text}\`);
            }
          } catch (err) {
            showMessage('error', \`网络错误: \${err.message}\`);
          }
        });
      }

      // --- 下载逻辑 ---
      const downloadForm = document.getElementById('download-form');
      if (downloadForm) {
        downloadForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          const homeworkName = document.getElementById('homeworkName').value;
          if (!homeworkName) {
            showMessage('error', '请输入作业名称');
            return;
          }
          
          showMessage('loading', \`正在生成 "\${homeworkName}" 的汇总表...\`);
          
          const url = \`/summary?homework=\${encodeURIComponent(homeworkName)}\`;
          
          try {
            // 注意: 我们要处理的是文件下载, 而不是 JSON
            const response = await fetch(url);
            
            if (response.ok) {
              // 检查返回的是否是 CSV
              const contentType = response.headers.get('Content-Type');
              if (contentType && contentType.includes('text/csv')) {
                // 成功, 触发下载
                const blob = await response.blob();
                const disposition = response.headers.get('Content-Disposition');
                let filename = \`\${homeworkName}_summary.csv\`;
                if (disposition && disposition.includes('filename=')) {
                  filename = disposition.split('filename=')[1].replace(/"/g, '');
                }
                
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                
                showMessage('success', \`"\${filename}" 已开始下载。\`);
              } else {
                 // 可能是 JSON 错误
                 const text = await response.text();
                 showMessage('error', \`服务器返回非CSV内容: \${text}\`);
              }
            } else {
              // 404 或 500 错误
              const text = await response.text();
              showMessage('error', \`无法下载: \${text}\`);
            }
          } catch (err) {
            showMessage('error', \`网络错误: \${err.message}\`);
          }
        });
      }
    </script>
  `;

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${appTitle}</title>
      ${commonStyles}
    </head>
    <body>
      <main>
        ${bodyContent}
      </main>
      ${scripts}
    </body>
    </html>
  `;
}

