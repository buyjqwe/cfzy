/**
 * Cloudflare Worker - AI 作业批改平台 (V4 - 分片代理 + 远程解压)
 *
 * 架构:
 * 1. 认证: 教师使用邮箱验证码登录 (无白名单限制)，会话存储在 Cookie 中。
 * 2. 大文件上传:
 * - 前端 JS 将大文件 (>100MB) 切片 (每片 10MB)。
 * - 前端请求 Worker 创建一个 OneDrive 上传会话 (Upload Session)。
 * - 前端将每个分片 POST 到 Worker (`/upload-chunk`)。
 * - Worker 立即将该分片 PUT 到 OneDrive 的上传会话 URL，实现“流式代理”。
 * - 此方法绕过了 Worker 的 100MB 入口限制 和 128MB 内存限制。
 * 3. AI 批改 (后台):
 * - Worker 命令 OneDrive 远程解压 (`/extract`) ZIP 包。
 * - Worker 遍历解压后的文件夹，逐个下载学生文件。
 * - Worker 将学生文件发送给 Gemini API。
 * - Worker 将批改报告 (report.json) 上传回学生目录。
 * 4. 结果下载:
 * - 教师可以下载一个按需生成的 CSV 成绩总表 (`/summary`)。
 */

import { GoogleGenerativeAI } from '@google/generative-ai';

// --- 常量 ---
const TOKEN_COOKIE = 'auth_token';
const TOKEN_EXPIRATION_SECONDS = 7 * 24 * 60 * 60; // 7 天
// OneDrive 分片上传大小 (10 MiB)。必须小于 60 MiB。
const CHUNK_SIZE = 10 * 1024 * 1024;

/**
 * 主入口点
 */
export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      // 1. 公开路由 (登录/验证)
      if (url.pathname === '/login' && request.method === 'POST') {
        return await handleLogin(request, env);
      }
      if (url.pathname === '/verify' && request.method === 'POST') {
        return await handleVerify(request, env, ctx);
      }

      // 2. 检查所有受保护路由的会话
      const session = await getSession(request, env);
      if (!session) {
        // 如果是 API 请求 (如 /upload)，返回 401
        if (url.pathname.startsWith('/api/')) {
          return jsonResponse({ error: 'Unauthorized' }, 401);
        }
        // 否则，重定向到登录页面
        return new Response(null, {
          status: 302,
          headers: { Location: '/login-page' },
        });
      }

      // 3. 受保护的路由
      switch (url.pathname) {
        case '/':
        case '/upload':
          return new Response(getHtmlPage(env.APP_TITLE, session.email, 'upload'), {
            headers: { 'Content-Type': 'text/html; charset=utf-8' },
          });
        case '/login-page':
          return new Response(getHtmlPage(env.APP_TITLE, null, 'login'), {
            headers: { 'Content-Type': 'text/html; charset=utf-8' },
          });
        case '/summary':
          return new Response(getHtmlPage(env.APP_TITLE, session.email, 'summary'), {
            headers: { 'Content-Type': 'text/html; charset=utf-8' },
          });
        case '/logout':
          return handleLogout(request, env);

        // --- API 路由 (由前端 JS 调用) ---
        case '/api/create-upload-session':
          if (request.method === 'POST') {
            return await handleCreateUploadSession(request, env, session);
          }
          break;
        case '/api/upload-chunk':
          if (request.method === 'POST') {
            return await handleUploadChunk(request, env, session);
          }
          break;
        case '/api/complete-upload':
          if (request.method === 'POST') {
            ctx.waitUntil(handleBackgroundProcessing(request, env, session));
            return jsonResponse({ success: true, message: '文件已接收，正在启动后台处理...' });
          }
          break;
        case '/api/download-summary':
          if (request.method === 'POST') {
            return await handleSummaryDownload(request, env, session);
          }
          break;
      }

      // 默认重定向到主页
      if (url.pathname !== '/login-page') {
        return new Response(null, {
          status: 302,
          headers: { Location: '/' },
        });
      }

      return new Response('Not Found', { status: 404 });
    } catch (err) {
      console.error('Fetch handler error:', err);
      return jsonResponse({ error: err.message || 'Internal Server Error' }, 500);
    }
  },
};

// =============================================
// 认证和会话
// =============================================

/**
 * 1. 处理登录请求 (POST /login)
 * - 生成验证码，存入 KV，发送邮件
 */
async function handleLogin(request, env) {
  let email;
  try {
    const { email: reqEmail } = await request.json();
    email = reqEmail;
  } catch (e) {
    return jsonResponse({ error: 'Invalid JSON request' }, 400);
  }

  if (!email || !email.includes('@')) {
    return jsonResponse({ error: 'Invalid email format' }, 400);
  }

  // [!!] 按照用户要求，移除了白名单检查

  try {
    const code = Math.floor(100000 + Math.random() * 900000).toString();
    const kvKey = `code:${email}`;

    // 将验证码存入 KV，有效期 5 分钟
    await env.AUTH_KV.put(kvKey, code, { expirationTtl: 300 });

    // 发送邮件
    await sendVerificationEmail(email, code, env);

    return jsonResponse({ success: true, message: `Verification code sent to ${email}` });
  } catch (err) {
    console.error('Login error:', err);
    return jsonResponse({ error: `Failed to send email: ${err.message}` }, 500);
  }
}

/**
 * 2. 处理验证码 (POST /verify)
 * - 验证 KV 中的验证码，创建会话，设置 Cookie
 */
async function handleVerify(request, env, ctx) {
  let email, code;
  try {
    const { email: reqEmail, code: reqCode } = await request.json();
    email = reqEmail;
    code = reqCode;
  } catch (e) {
    return jsonResponse({ error: 'Invalid JSON request' }, 400);
  }

  if (!email || !code) {
    return jsonResponse({ error: 'Email and code are required' }, 400);
  }

  const kvKey = `code:${email}`;
  const storedCode = await env.AUTH_KV.get(kvKey);

  if (!storedCode) {
    return jsonResponse({ error: 'Verification code expired or invalid' }, 401);
  }

  if (storedCode !== code) {
    return jsonResponse({ error: 'Verification code does not match' }, 401);
  }

  // 验证成功，删除 KV 中的验证码
  ctx.waitUntil(env.AUTH_KV.delete(kvKey));

  // 创建会话
  const sessionToken = crypto.randomUUID();
  const sessionKey = `session:${sessionToken}`;
  const sessionData = JSON.stringify({ email });

  await env.AUTH_KV.put(sessionKey, sessionData, {
    expirationTtl: TOKEN_EXPIRATION_SECONDS,
  });

  // 设置安全 Cookie 并重定向
  const headers = new Headers();
  headers.append(
    'Set-Cookie',
    `${TOKEN_COOKIE}=${sessionToken}; HttpOnly; Secure; Path=/; Max-Age=${TOKEN_EXPIRATION_SECONDS}`
  );
  // [!!] BUG FIX: 修复 "Unexpected token 'E'"
  // 不再发送 302，而是发送 JSON 响应，让前端 JS 处理跳转
  return jsonResponse({ success: true, redirect: '/' });
}

/**
 * 3. 处理登出 (GET /logout)
 * - 删除 KV 中的会话并清除 Cookie
 */
async function handleLogout(request, env) {
  const token = getCookie(request, TOKEN_COOKIE);
  if (token) {
    await env.AUTH_KV.delete(`session:${token}`);
  }

  // 清除 Cookie 并重定向到登录页
  const headers = new Headers();
  headers.append('Set-Cookie', `${TOKEN_COOKIE}=; HttpOnly; Secure; Path=/; Max-Age=0`);
  headers.append('Location', '/login-page');
  return new Response(null, { status: 302, headers });
}

/**
 * 4. 检查请求是否包含有效会话
 */
async function getSession(request, env) {
  const token = getCookie(request, TOKEN_COOKIE);
  if (!token) {
    return null;
  }

  const sessionKey = `session:${token}`;
  const data = await env.AUTH_KV.get(sessionKey);

  if (!data) {
    return null;
  }

  return JSON.parse(data); // { email: "..." }
}

// =============================================
// Microsoft Graph API (邮件 + OneDrive)
// =============================================

/**
 * 获取 Microsoft Graph API 访问令牌
 * (使用 Client Credentials Flow)
 */
async function getMSGraphToken(env) {
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const response = await fetch(url, {
    method: 'POST',
    body: body,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`MS Token Error: ${error}`);
  }

  const data = await response.json();
  return data.access_token;
}

/**
 * (Graph API) 发送邮件
 */
async function sendVerificationEmail(toEmail, code, env) {
  const token = await getMSGraphToken(env);
  const url = `https://graph.microsoft.com/v1.0/users/${env.MS_USER_ID}/sendMail`;

  const emailBody = {
    message: {
      subject: `[${env.APP_TITLE}] 您的登录验证码`,
      body: {
        contentType: 'HTML',
        content: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2>您的登录验证码</h2>
            <p>您好，</p>
            <p>您正在登录 ${env.APP_TITLE}。您的验证码是：</p>
            <h1 style="font-size: 3em; letter-spacing: 5px; margin: 20px 0; color: #333;">
              ${code}
            </h1>
            <p>此验证码将在 5 分钟后过期。</p>
            <p>如果您没有请求此验证码，请忽略此邮件。</p>
          </div>
        `,
      },
      toRecipients: [
        {
          emailAddress: {
            address: toEmail,
          },
        },
      ],
    },
    saveToSentItems: 'true',
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(emailBody),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Send mail error:', errorText);
    throw new Error(`Failed to send mail: ${response.statusText}`);
  }
}

/**
 * (Graph API) 为大文件创建 OneDrive 上传会话
 * @param {string} accessToken
 * @param {string} pathOnOneDrive - e.g., "Apps/HomeworkGrader/homework.zip"
 * @param {string} MS_USER_ID - 环境变量
 * @returns {object} { uploadUrl: "...", expirationDateTime: "..." }
 */
async function createUploadSession(accessToken, pathOnOneDrive, MS_USER_ID) {
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/root:/${pathOnOneDrive}:/createUploadSession`;

  const response = await fetch(graphUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      item: {
        '@microsoft.graph.conflictBehavior': 'replace',
      },
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Create Upload Session failed: ${error}`);
  }
  return await response.json();
}

/**
 * (Graph API) 命令 OneDrive 远程解压 ZIP 文件
 * @param {string} accessToken
 * @param {string} zipFileItemId - OneDrive 中 zip 文件的 Item ID
 * @param {string} destinationFolderItemId - 解压目标文件夹的 Item ID
 * @param {string} MS_USER_ID - 环境变量
 * @returns {object} 异步操作的监控 URL
 */
async function extractZipOnOneDrive(accessToken, zipFileItemId, destinationFolderItemId, MS_USER_ID) {
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/items/${zipFileItemId}/extract`;

  const response = await fetch(graphUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      parentReference: {
        id: destinationFolderItemId,
      },
    }),
  });

  // 202 Accepted 表示操作已开始
  if (response.status === 202) {
    // 返回用于轮询状态的 URL
    return response.headers.get('Location');
  } else {
    const error = await response.text();
    throw new Error(`OneDrive Extract command failed: ${error}`);
  }
}

/**
 * (Graph API) 轮询 OneDrive 异步操作 (如解压) 的状态
 * @param {string} accessToken
 * @param {string} monitorUrl - 从 extractZipOnOneDrive 返回的 URL
 */
async function pollOperationStatus(accessToken, monitorUrl) {
  let status = 'inProgress';
  while (status === 'inProgress' || status === 'notStarted' || status === 'running') {
    await new Promise((resolve) => setTimeout(resolve, 3000)); // 等待 3 秒

    const response = await fetch(monitorUrl, {
      method: 'GET',
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!response.ok) {
      throw new Error('Failed to poll operation status');
    }
    const result = await response.json();
    status = result.status;
    console.log(`[OneDrive Extract] Status: ${status}`);

    if (status === 'completed' || status === 'succeeded') {
      return true;
    }
    if (status === 'failed') {
      throw new Error(`OneDrive Extract operation failed: ${result.error?.message}`);
    }
  }
}

/**
 * (Graph API) 获取 OneDrive 项的 Item ID (通过路径)
 * @param {string} accessToken
 * @param {string} pathOnOneDrive - e.g., "Apps/HomeworkGrader/homework.zip"
 * @param {string} MS_USER_ID - 环境变量
 * @returns {string} Item ID
 */
async function getItemIdByPath(accessToken, pathOnOneDrive, MS_USER_ID) {
  // 确保路径以 /drive/root:/ 开头，并且正确编码
  const encodedPath = encodeURI(pathOnOneDrive).replace(/'/g, "''").replace(/%/g, '%25');
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/root:/${encodedPath}`;

  const response = await fetch(graphUrl, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Get Item ID failed for path ${pathOnOneDrive}: ${error}`);
  }
  const data = await response.json();
  return data.id;
}

/**
 * (Graph API) 获取文件夹中的所有子项 (文件/文件夹)
 * @param {string} accessToken
 * @param {string} folderItemId - 文件夹的 Item ID
 * @param {string} MS_USER_ID - 环境变量
 * @returns {array} 子项列表
 */
async function listChildrenInFolder(accessToken, folderItemId, MS_USER_ID) {
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/items/${folderItemId}/children`;

  const response = await fetch(graphUrl, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`List children failed: ${error}`);
  }
  const data = await response.json();
  return data.value; // [{ id: "...", name: "...", file: {}, folder: {} }, ...]
}

/**
 * (Graph API) 下载小文件内容 (< 4MB)
 * @param {string} accessToken
 * @param {string} fileItemId - 文件的 Item ID
 * @param {string} MS_USER_ID - 环境变量
 * @returns {Promise<ArrayBuffer>} 文件内容
 */
async function downloadSmallFile(accessToken, fileItemId, MS_USER_ID) {
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/items/${fileItemId}/content`;

  const response = await fetch(graphUrl, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Download file failed: ${error}`);
  }
  return response.arrayBuffer();
}

/**
 * (Graph API) 上传小文件 (如 report.json)
 * @param {string} accessToken
 * @param {string} pathOnOneDrive - e.g., "Apps/HomeworkGrader/hw1/studentA/report.json"
 * @param {string} content - 文件内容 (字符串)
 * @param {string} MS_USER_ID - 环境变量
 */
async function uploadSmallFile(accessToken, pathOnOneDrive, content, MS_USER_ID) {
  const encodedPath = encodeURI(pathOnOneDrive).replace(/'/g, "''").replace(/%/g, '%25');
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${MS_USER_ID}/drive/root:/${encodedPath}:/content`;

  const response = await fetch(graphUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: content,
  });

  if (!response.ok && response.status !== 201) {
    const error = await response.text();
    throw new Error(`Upload small file failed: ${error}`);
  }
  return await response.json();
}

// =============================================
// V4 - 分片上传 (API 路由)
// =============================================

/**
 * API (POST /api/create-upload-session)
 * - 前端请求开始上传，Worker 从 OneDrive 获取上传 URL
 */
async function handleCreateUploadSession(request, env, session) {
  const { filename } = await request.json();
  if (!filename) {
    return jsonResponse({ error: 'Filename is required' }, 400);
  }

  // 规范化文件名，移除路径字符
  const safeFilename = filename.split(/[\/\\]/).pop();
  const homeworkName = safeFilename.replace(/\.zip$/i, ''); // "my-homework.zip" -> "my-homework"
  const pathOnOneDrive = `${env.MS_ONEDRIVE_BASE_PATH}/${safeFilename}`;

  try {
    const token = await getMSGraphToken(env);
    const sessionData = await createUploadSession(token, pathOnOneDrive, env.MS_USER_ID);

    // 将上传会话 URL 和作业元数据存入 KV，供后续分片和完成时使用
    // 使用会话 token 作为键，确保只有该用户能操作
    const sessionKey = `upload:${session.email}:${safeFilename}`;
    const kvData = {
      uploadUrl: sessionData.uploadUrl,
      expiration: sessionData.expirationDateTime,
      homeworkName: homeworkName,
      zipFilePath: pathOnOneDrive,
    };
    await env.AUTH_KV.put(sessionKey, JSON.stringify(kvData), {
      expirationTtl: 60 * 60 * 2, // 2 小时
    });

    console.log(`[Upload Session] Created for ${safeFilename}`);
    return jsonResponse({
      success: true,
      uploadUrl: `/api/upload-chunk?sessionKey=${encodeURIComponent(sessionKey)}`, // 前端应 POST 到的 Worker 代理 URL
      chunkSize: CHUNK_SIZE,
    });
  } catch (err) {
    console.error('Create Session Error:', err);
    return jsonResponse({ error: `Failed to create upload session: ${err.message}` }, 500);
  }
}

/**
 * API (POST /api/upload-chunk)
 * - Worker 作为代理，接收前端发来的分片，并立即将其 PUT 到 OneDrive
 */
async function handleUploadChunk(request, env, session) {
  const url = new URL(request.url);
  const sessionKey = url.searchParams.get('sessionKey');
  const contentRange = request.headers.get('Content-Range');
  const contentLength = request.headers.get('Content-Length');

  if (!sessionKey || !contentRange || !contentLength) {
    return jsonResponse({ error: 'Missing sessionKey, Content-Range, or Content-Length' }, 400);
  }

  // 验证 sessionKey 是否属于当前登录用户
  if (!sessionKey.startsWith(`upload:${session.email}:`)) {
    return jsonResponse({ error: 'Session key mismatch' }, 403);
  }

  const kvDataStr = await env.AUTH_KV.get(sessionKey);
  if (!kvDataStr) {
    return jsonResponse({ error: 'Upload session expired or not found' }, 404);
  }

  const kvData = JSON.parse(kvDataStr);
  const oneDriveUploadUrl = kvData.uploadUrl;

  try {
    // 流式代理：将请求体 (分片) 直接 PUT 到 OneDrive
    const response = await fetch(oneDriveUploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Range': contentRange,
        'Content-Length': contentLength,
      },
      body: request.body,
    });

    if (!response.ok && response.status !== 201 && response.status !== 202) {
      const error = await response.text();
      throw new Error(`OneDrive Chunk Upload failed: ${error}`);
    }

    const data = await response.json();

    // 返回 OneDrive 的响应 (通常包含 nextExpectedRanges)
    return jsonResponse(data, response.status);
  } catch (err) {
    console.error('Upload Chunk Error:', err);
    return jsonResponse({ error: `Failed to upload chunk: ${err.message}` }, 500);
  }
}

// =============================================
// V4 - 后台处理 (远程解压 + Gemini)
// =============================================

/**
 * API (POST /api/complete-upload)
 * - 前端通知 Worker 所有分片已完成
 * - Worker 在后台 (waitUntil) 启动远程解压和 AI 批改
 */
async function handleBackgroundProcessing(request, env, session) {
  const { sessionKey } = await request.json();
  if (!sessionKey || !sessionKey.startsWith(`upload:${session.email}:`)) {
    throw new Error('Invalid or missing sessionKey');
  }

  const kvDataStr = await env.AUTH_KV.get(sessionKey);
  if (!kvDataStr) {
    throw new Error('Upload session not found in KV');
  }
  // 删除临时的上传会话
  await env.AUTH_KV.delete(sessionKey);

  const kvData = JSON.parse(kvDataStr);
  const { homeworkName, zipFilePath } = kvData;
  const basePath = env.MS_ONEDRIVE_BASE_PATH;
  const homeworkPath = `${basePath}/${homeworkName}`; // "Apps/HomeworkGrader/my-homework"

  console.log(`[BG Process Start] Homework: ${homeworkName}`);

  try {
    const token = await getMSGraphToken(env);
    const MS_USER_ID = env.MS_USER_ID; // 方便后续函数调用

    // 1. 获取 ZIP 文件的 Item ID
    console.log(`[BG Process] Getting ZIP Item ID for: ${zipFilePath}`);
    const zipFileId = await getItemIdByPath(token, zipFilePath, MS_USER_ID);

    // 2. 获取目标文件夹 (homeworkPath) 的 Item ID
    // (Graph API 不能在解压时自动创建父目录，我们必须先获取父目录 ID)
    console.log(`[BG Process] Getting Base Folder ID for: ${homeworkPath}`);
    const homeworkFolderId = await getItemIdByPath(token, homeworkPath, MS_USER_ID);

    // 3. 命令 OneDrive 远程解压
    console.log(`[BG Process] Sending Extract command...`);
    const monitorUrl = await extractZipOnOneDrive(token, zipFileId, homeworkFolderId, MS_USER_ID);

    // 4. 轮询解压状态
    console.log(`[BG Process] Polling extract status...`);
    await pollOperationStatus(token, monitorUrl);
    console.log(`[BG Process] Extract complete!`);

    // 5. 初始化 Gemini
    const genAI = new GoogleGenerativeAI(env.GOOGLE_API_KEY);
    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' }); // 假设使用 1.5 Flash

    // 6. 遍历解压后的学生文件夹
    // (解压后，文件结构应为: .../my-homework/学生A/file1.pdf, .../my-homework/学生B/file2.txt)
    const studentFolders = await listChildrenInFolder(token, homeworkFolderId, MS_USER_ID);

    for (const studentFolder of studentFolders) {
      // 确保是文件夹
      if (!studentFolder.folder) continue;

      const studentName = studentFolder.name;
      const studentFolderId = studentFolder.id;
      console.log(`[BG Process] Processing student: ${studentName}`);

      try {
        const studentFiles = await listChildrenInFolder(token, studentFolderId, MS_USER_ID);
        const promptParts = [
          getGradingPrompt(),
          `# 学生: ${studentName}`,
          '# 学生提交的文件内容如下:',
        ];

        let fileCount = 0;
        for (const file of studentFiles) {
          if (!file.file) continue; // 忽略子文件夹

          // 避免批改自己的报告
          if (file.name === 'report.json' || file.name === 'error_report.json') {
            continue;
          }

          console.log(`[BG Process]   Reading file: ${file.name}`);
          const fileId = file.id;
          const fileContent = await downloadSmallFile(token, fileId, MS_USER_ID);

          promptParts.push(`--- FILE: ${file.name} ---`);

          // 将 ArrayBuffer 转换为 Base64 字符串并指定 MIME 类型
          const base64Data = arrayBufferToBase64(fileContent);
          const mimeType = file.file.mimeType || 'application/octet-stream';

          promptParts.push({
            inlineData: {
              data: base64Data,
              mimeType: mimeType,
            },
          });
          fileCount++;
        }

        if (fileCount === 0) {
          console.log(`[BG Process]   No files found for ${studentName}. Skipping.`);
          continue;
        }

        // 7. 调用 Gemini API
        console.log(`[BG Process]   Calling Gemini for ${studentName}...`);
        const result = await model.generateContent({ contents: [{ role: 'user', parts: promptParts }] });
        const responseText = result.response.text();

        // 8. 上传批改报告
        const reportPath = `${homeworkPath}/${studentName}/report.json`;
        console.log(`[BG Process]   Uploading report for ${studentName} to ${reportPath}`);

        // 尝试解析 AI 的 JSON，如果失败，也上传原始文本
        let reportContent;
        try {
          // 确保 AI 返回的是有效的 JSON (更稳健的解析)
          const jsonMatch = responseText.match(/\{[\s\S]*\}/);
          if (!jsonMatch) throw new Error('No JSON object found in AI response');
          
          const jsonText = jsonMatch[0];
          JSON.parse(jsonText); // 验证
          reportContent = jsonText;
        } catch (e) {
          console.error(`[BG Process] Gemini response for ${studentName} was not valid JSON. Saving raw text. Error: ${e.message}`);
          reportContent = JSON.stringify({
            error: 'AI response was not valid JSON',
            raw_response: responseText,
            grade: 0,
            feedback: 'AI 批改失败，返回的不是有效 JSON。请联系管理员。',
          });
        }

        await uploadSmallFile(token, reportPath, reportContent, MS_USER_ID);
      } catch (studentErr) {
        console.error(`[BG Process] Failed to process student ${studentName}: ${studentErr.message}`);
        // 尝试上传一个错误报告
        try {
          const errorReportPath = `${homeworkPath}/${studentName}/error_report.json`;
          await uploadSmallFile(
            token,
            errorReportPath,
            JSON.stringify({
              error: `Failed to process this student: ${studentErr.message}`,
              stack: studentErr.stack,
            }),
            MS_USER_ID
          );
        } catch (reportErr) {
          console.error(`[BG Process] Failed to upload error report for ${studentName}: ${reportErr.message}`);
        }
      }
    }
    console.log(`[BG Process Finish] Homework: ${homeworkName}`);
  } catch (err) {
    console.error(`[BG Process Fatal Error] ${err.message}`);
    // 可以在此处添加逻辑，例如上传一个总的错误文件到作业根目录
  }
}

// =============================================
// V4 - 成绩汇总 (API 路由)
// =============================================

/**
 * API (POST /api/download-summary)
 * - 教师请求下载指定作业的 CSV 成绩总表
 */
async function handleSummaryDownload(request, env, session) {
  const { homeworkName } = await request.json();
  if (!homeworkName) {
    return jsonResponse({ error: 'Homework name is required' }, 400);
  }

  const homeworkPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}`;
  const results = [];

  try {
    const token = await getMSGraphToken(env);
    const MS_USER_ID = env.MS_USER_ID;

    // 1. 获取作业文件夹 ID
    const homeworkFolderId = await getItemIdByPath(token, homeworkPath, MS_USER_ID);

    // 2. 遍历学生文件夹
    const studentFolders = await listChildrenInFolder(token, homeworkFolderId, MS_USER_ID);

    for (const studentFolder of studentFolders) {
      if (!studentFolder.folder) continue;
      const studentName = studentFolder.name;

      try {
        // 3. 查找 report.json
        const studentFiles = await listChildrenInFolder(token, studentFolder.id, MS_USER_ID);
        const reportFile = studentFiles.find((f) => f.name === 'report.json');

        if (reportFile) {
          // 4. 下载并解析 report.json
          const reportContentBuffer = await downloadSmallFile(token, reportFile.id, MS_USER_ID);
          const reportJsonStr = new TextDecoder().decode(reportContentBuffer);
          const reportData = JSON.parse(reportJsonStr);

          results.push({
            student: studentName,
            grade: reportData.grade || 'N/A',
            feedback: reportData.feedback || 'N/A',
            error: reportData.error || '',
          });
        } else {
          results.push({
            student: studentName,
            grade: '批改未完成',
            feedback: '',
            error: 'report.json not found',
          });
        }
      } catch (studentErr) {
        results.push({
          student: studentName,
          grade: '错误',
          feedback: '',
          error: `Failed to parse report: ${studentErr.message}`,
        });
      }
    }

    // 5. 生成 CSV
    const csv = generateCsv(results);
    const filename = `summary_${homeworkName.replace(/[^a-z0-9]/gi, '_')}.csv`;

    return new Response(csv, {
      headers: {
        'Content-Type': 'text/csv; charset=utf-8-sig', // utf-8-sig 确保 Excel 正确打开
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    console.error('Summary Download Error:', err);
    return jsonResponse({ error: `Failed to generate summary: ${err.message}` }, 500);
  }
}

// =============================================
// 辅助函数
// =============================================

/**
 * (辅助) 生成 AI 批改指令
 */
function getGradingPrompt() {
  return `
# 角色
你是一位严格、公正的 AI 教学助手。

# 任务
根据学生提交的多个文件（可能是 PDF, TXT, MD, PY, IPYNB, 图片等），综合评估他们的作业完成情况。

# 核心指令
1.  **综合分析**：你必须阅读和理解所有提供的文件，学生的答案可能分散在多个文件中。
2.  **严格评分**：根据作业的隐含要求（例如代码的正确性、分析的深度、报告的完整性）给出一个 0-100 的分数。
3.  **提供反馈**：提供清晰、有建设性的评语，指出学生的优点和主要存在的问题。

# 输出格式
你必须严格按照以下 JSON 格式返回你的批改结果，不要包含任何 markdown 标记 (如 \`\`\`json)。

{
  "grade": 85,
  "feedback": "学生A，你对基础概念的理解很好，PDF 报告中的分析很到位。但提交的代码（main.py）在处理边界条件时存在一个逻辑错误，导致部分测试无法通过。请修正 xxx 部分。"
}
`;
}

/**
 * (辅助) 将批改结果数组转换为 CSV 字符串
 */
function generateCsv(results) {
  const header = ['Student', 'Grade', 'Feedback', 'Error'];
  const rows = results.map((r) =>
    [
      `"${r.student.replace(/"/g, '""')}"`,
      r.grade,
      `"${(r.feedback || '').replace(/"/g, '""').replace(/\n/g, ' ')}"`, // 移除反馈中的换行符
      `"${(r.error || '').replace(/"/g, '""')}"`,
    ].join(',')
  );
  return [header.join(','), ...rows].join('\n');
}

/**
 * (辅助) 解析 Cookie
 */
function getCookie(request, name) {
  const cookieHeader = request.headers.get('Cookie');
  if (!cookieHeader) return null;
  const cookies = cookieHeader.split(';');
  for (const cookie of cookies) {
    const [key, value] = cookie.trim().split('=');
    if (key === name) {
      return value;
    }
  }
  return null;
}

/**
 * (辅助) 返回 JSON 响应
 */
function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/**
 * (辅助) ArrayBuffer to Base64
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

// =============================================
// V4 - 前端 HTML / JS 页面
// =============================================

function getHtmlPage(appTitle, userEmail, mode = 'login') {
  const isLoggedIn = Boolean(userEmail);
  const title = `${appTitle} - ${isLoggedIn ? userEmail : '请登录'}`;

  // 导航栏
  const nav = isLoggedIn
    ? `
    <nav class="flex justify-between items-center p-4 bg-gray-800 text-white shadow-md">
      <div>
        <a href="/" class="text-lg font-semibold hover:text-blue-300 ${mode === 'upload' ? 'text-blue-400' : ''}">上传作业</a>
        <a href="/summary" class="ml-4 text-lg font-semibold hover:text-blue-300 ${mode === 'summary' ? 'text-blue-400' : ''}">下载总表</a>
      </div>
      <div class="flex items-center">
        <span class="mr-4 text-gray-300">${userEmail}</span>
        <a href="/logout" class="px-3 py-1 bg-red-600 rounded hover:bg-red-700">登出</a>
      </div>
    </nav>
  `
    : `
    <nav class="flex justify-between items-center p-4 bg-gray-800 text-white shadow-md">
      <h1 class="text-xl font-bold">${appTitle}</h1>
      <span class="text-gray-400">请登录</span>
    </nav>
  `;

  // 页面内容
  let content = '';
  if (mode === 'login') {
    content = `
    <div class="w-full max-w-md">
      <h2 class="text-2xl font-bold text-center mb-6 text-gray-800">教师登录</h2>
      
      <!-- Step 1: Email -->
      <form id="login-form" class="bg-white p-8 rounded-lg shadow-lg">
        <div class="mb-4">
          <label for="email" class="block text-sm font-medium text-gray-700 mb-2">教师邮箱</label>
          <input type="email" id="email" name="email" required
                 class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
        </div>
        <button type="submit" id="login-btn"
                class="w-full bg-blue-600 text-white py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition duration-300">
          发送验证码
        </button>
      </form>

      <!-- Step 2: Code -->
      <form id="verify-form" class="hidden bg-white p-8 rounded-lg shadow-lg">
        <p class="text-sm text-gray-600 mb-4">验证码已发送至 <strong id="email-display"></strong></p>
        <div class="mb-4">
          <label for="code" class="block text-sm font-medium text-gray-700 mb-2">6 位验证码</label>
          <input type="text" id="code" name="code" required minlength="6" maxlength="6"
                 class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 tracking-widest text-center">
        </div>
        <button type="submit" id="verify-btn"
                class="w-full bg-green-600 text-white py-2 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-opacity-50 transition duration-300">
          登录
        </button>
        <button type="button" id="back-btn" class="w-full text-center text-sm text-gray-600 mt-4 hover:underline">返回</button>
      </form>
      <div id="message-box" class="mt-4 text-center"></div>
    </div>
    `;
  } else if (mode === 'upload') {
    content = `
    <div class="w-full max-w-2xl">
      <h2 class="text-2xl font-bold text-center mb-6 text-gray-800">上传作业压缩包 (支持大文件)</h2>
      <div class="bg-white p-8 rounded-lg shadow-lg">
        <div class="mb-4">
          <label for="file-input" class="block text-sm font-medium text-gray-700 mb-2">选择作业 .zip 文件</label>
          <input type="file" id="file-input" name="file" accept=".zip"
                 class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
        </div>
        <button id="upload-btn"
                class="w-full bg-blue-600 text-white py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition duration-300">
          上传并开始后台批改
        </button>
        <div id="upload-progress-bar" class="w-full bg-gray-200 rounded-full h-4 mt-4 hidden">
          <div id="upload-progress" class="bg-blue-500 h-4 rounded-full" style="width: 0%"></div>
        </div>
        <div id="message-box" class="mt-4 text-center"></div>
      </div>
    </div>
    `;
  } else if (mode === 'summary') {
    content = `
    <div class="w-full max-w-2xl">
      <h2 class="text-2xl font-bold text-center mb-6 text-gray-800">下载成绩总表 (CSV)</h2>
      <form id="summary-form" class="bg-white p-8 rounded-lg shadow-lg">
        <div class="mb-4">
          <label for="homework-name" class="block text-sm font-medium text-gray-700 mb-2">
            作业名称
            <span class="text-xs text-gray-500">(与上传的 .zip 文件名一致，不含 .zip 后缀)</span>
          </label>
          <input type="text" id="homework-name" name="homeworkName" required
                 class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                 placeholder="例如: 22计网1-GET+HEAD(附件)">
        </div>
        <button type="submit" id="summary-btn"
                class="w-full bg-green-600 text-white py-2 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-opacity-50 transition duration-300">
          下载成绩总表 (.csv)
        </button>
        <div id="message-box" class="mt-4 text-center"></div>
      </form>
    </div>
    `;
  }

  // 返回完整 HTML
  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${title}</title>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-100 min-h-screen">
      ${nav}
      <main class="flex items-center justify-center p-6" style="min-height: calc(100vh - 68px);">
        ${content}
      </main>

      <script>
        // 辅助函数
        const msgBox = document.getElementById('message-box');
        const showMessage = (message, isError = false) => {
          if (!msgBox) return;
          msgBox.textContent = message;
          msgBox.className = isError 
            ? 'mt-4 text-center text-red-600 p-2 bg-red-100 rounded' 
            : 'mt-4 text-center text-green-600 p-2 bg-green-100 rounded';
        };

        // --- 登录页面逻辑 ---
        if (document.getElementById('login-form')) {
          const loginForm = document.getElementById('login-form');
          const verifyForm = document.getElementById('verify-form');
          const loginBtn = document.getElementById('login-btn');
          const emailInput = document.getElementById('email');
          const emailDisplay = document.getElementById('email-display');
          const backBtn = document.getElementById('back-btn');

          loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const email = emailInput.value;
            loginBtn.disabled = true;
            loginBtn.textContent = '发送中...';
            showMessage('正在发送验证码...', false);

            try {
              const res = await fetch('/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email })
              });
              const data = await res.json();
              if (!res.ok) throw new Error(data.error || '请求失败');
              
              showMessage(data.message, false);
              emailDisplay.textContent = email;
              loginForm.classList.add('hidden');
              verifyForm.classList.remove('hidden');
            } catch (err) {
              showMessage('错误: ' + err.message, true);
            } finally {
              loginBtn.disabled = false;
              loginBtn.textContent = '发送验证码';
            }
          });

          verifyForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const verifyBtn = document.getElementById('verify-btn');
            verifyBtn.disabled = true;
            verifyBtn.textContent = '登录中...';
            
            try {
              const res = await fetch('/verify', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: emailInput.value, code: document.getElementById('code').value })
              });
              
              const data = await res.json();
              if (!res.ok) {
                throw new Error(data.error || '验证失败');
              }

              // [!!] 修复: 检查 JSON 响应中的 redirect
              if (data.success && data.redirect) {
                 showMessage('登录成功！正在跳转...', false);
                 window.location.href = data.redirect; // 手动跳转
              } else {
                 throw new Error('未知的登录响应');
              }

            } catch (err) {
              showMessage('错误: ' + err.message, true);
              verifyBtn.disabled = false;
              verifyBtn.textContent = '登录';
            }
          });
          
          backBtn.addEventListener('click', () => {
             loginForm.classList.remove('hidden');
             verifyForm.classList.add('hidden');
             showMessage('');
          });
        }

        // --- V4 - 分片上传逻辑 ---
        if (document.getElementById('upload-btn')) {
          const uploadBtn = document.getElementById('upload-btn');
          const fileInput = document.getElementById('file-input');
          const progressBar = document.getElementById('upload-progress-bar');
          const progress = document.getElementById('upload-progress');

          uploadBtn.addEventListener('click', async () => {
            const file = fileInput.files[0];
            if (!file) {
              showMessage('请先选择一个 .zip 文件。', true);
              return;
            }

            uploadBtn.disabled = true;
            uploadBtn.textContent = '正在初始化...';
            progressBar.classList.remove('hidden');
            progress.style.width = '0%';
            showMessage('正在请求 OneDrive 上传会话...', false);

            let sessionKey, chunkUploadUrl, chunkSize;
            
            try {
              // 1. 创建上传会话
              const sessionRes = await fetch('/api/create-upload-session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: file.name })
              });
              const sessionData = await sessionRes.json();
              if (!sessionRes.ok) throw new Error(sessionData.error || '创建会话失败');
              
              chunkUploadUrl = sessionData.uploadUrl; // 这是我们的 Worker 代理 URL
              sessionKey = new URLSearchParams(chunkUploadUrl.split('?')[1]).get('sessionKey');
              chunkSize = sessionData.chunkSize;
              
              showMessage('会话已创建，开始分片上传...', false);

              // 2. 循环分片上传
              let start = 0;
              while (start < file.size) {
                const end = Math.min(start + chunkSize, file.size);
                const chunk = file.slice(start, end);
                const progressPercent = Math.round((end / file.size) * 100);
                
                uploadBtn.textContent = \`上传中... (\${progressPercent}%)\`;
                progress.style.width = \`\${progressPercent}%\`;

                const contentRange = \`bytes \${start}-\${end - 1}/\${file.size}\`;
                
                const chunkRes = await fetch(chunkUploadUrl, {
                  method: 'POST',
                  headers: {
                    'Content-Range': contentRange,
                    'Content-Length': chunk.size,
                  },
                  body: chunk
                });

                if (!chunkRes.ok) {
                   const chunkError = await chunkRes.json();
                   throw new Error(chunkError.error?.message || '分片上传失败');
                }
                
                // OneDrive 返回 201 (完成) 或 202 (继续)
                const chunkData = await chunkRes.json();
                if (chunkData.nextExpectedRanges) {
                  // OneDrive 让我们从它建议的地方继续
                  start = parseInt(chunkData.nextExpectedRanges[0].split('-')[0], 10);
                } else {
                  // 201 Created (最后一块) 或 202 (但没有 nextExpectedRanges)
                  start = end;
                }
              }

              // 3. 通知 Worker 后台处理
              uploadBtn.textContent = '上传完成！正在启动后台批改...';
              showMessage('上传完成！正在启动后台批改...', false);
              
              const completeRes = await fetch('/api/complete-upload', {
                 method: 'POST',
                 headers: { 'Content-Type': 'application/json' },
                 body: JSON.stringify({ sessionKey: sessionKey })
              });
              
              const completeData = await completeRes.json();
              if (!completeRes.ok) throw new Error(completeData.error || '启动后台处理失败');

              showMessage(\`文件 "\${file.name}" 已成功上传并开始后台处理。这可能需要几分钟。您可以关闭此页面。\`, false);
              fileInput.value = ''; // 清空输入

            } catch (err) {
              showMessage('上传失败: ' + err.message, true);
              progressBar.classList.add('hidden');
            } finally {
              uploadBtn.disabled = false;
              uploadBtn.textContent = '上传并开始后台批改';
            }
          });
        }

        // --- 成绩汇总下载逻辑 ---
        if (document.getElementById('summary-form')) {
          const summaryForm = document.getElementById('summary-form');
          const summaryBtn = document.getElementById('summary-btn');

          summaryForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const homeworkName = document.getElementById('homework-name').value;
            if (!homeworkName) {
              showMessage('请输入作业名称。', true);
              return;
            }
            
            summaryBtn.disabled = true;
            summaryBtn.textContent = '正在生成...';
            showMessage('正在从 OneDrive 汇总数据...', false);

            try {
              const res = await fetch('/api/download-summary', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ homeworkName })
              });
              
              if (!res.ok) {
                 const data = await res.json();
                 throw new Error(data.error || '下载失败');
              }
              
              // 成功，触发浏览器下载
              const blob = await res.blob();
              const url = window.URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.style.display = 'none';
              a.href = url;
              // 从 Content-Disposition 获取文件名
              const disposition = res.headers.get('Content-Disposition');
              let filename = \`summary_\${homeworkName}.csv\`;
              if (disposition && disposition.indexOf('attachment') !== -1) {
                // [!!] BUG FIX: 2025-10-24
                // \2 必须转义为 \\2 才能在模板字符串中作为正则表达式的反向引用
                const filenameRegex = /filename[^;=\n]*=((['"]).*?\\2|[^;\n]*)/;
                const matches = filenameRegex.exec(disposition);
                if (matches != null && matches[1]) {
                  filename = matches[1].replace(/['"]/g, '');
                }
              }
              a.download = filename;
              document.body.appendChild(a);
              a.click();
              window.URL.revokeObjectURL(url);
              a.remove();
              showMessage('下载成功！', false);

            } catch (err) {
              showMessage('错误: ' + err.message, true);
            } finally {
              summaryBtn.disabled = false;
              summaryBtn.textContent = '下载成绩总表 (.csv)';
            }
          });
        }
      </script>
    </body>
    </html>
  `;
}

