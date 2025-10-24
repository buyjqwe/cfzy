/**
 * Cloudflare Worker - AI Homework Grader (v4 - Chunked Proxy + Remote Unzip)
 *
 * Features:
 * - 1. 教师使用邮箱验证码登录
 * - 2. 使用 Cookie 进行会话管理
 * - 3. Cloudflare KV 存储验证码、会话、上传会话
 * - 4. Microsoft Graph API 发送邮件
 * - 5. [!!] 分片代理上传 (Chunked Proxy Upload) 以支持 > 100MB 文件
 * - 6. [!!] OneDrive 远程解压 (/extract) 以避免 Worker 内存限制
 * - 7. 后台 AI 批改 (遍历远程文件夹)
 * - 8. 成绩汇总 (CSV) 下载
 */

import { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } from '@google/generative-ai';
// [!!] 'unzipit' 已被移除 (Removed 'unzipit')

// --- 会话和安全常量 (Session and Security Constants) ---
const SESSION_COOKIE_NAME = '__auth_session';
const SESSION_DURATION_SECONDS = 7 * 24 * 60 * 60; // 7 days
const CODE_TTL_SECONDS = 300; // 5 minutes
const UPLOAD_SESSION_TTL_SECONDS = 3600; // 1 hour for upload sessions

// --- 主路由 (Main Router) ---
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const path = url.pathname;
    const cookie = request.headers.get('Cookie') || '';
    const sessionToken = getCookie(cookie, SESSION_COOKIE_NAME);
    const session = await getSession(env.AUTH_KV, sessionToken);

    // 登录/验证路由 (Login/Verify Routes)
    if (path === '/login' && request.method === 'POST') {
      return handleLogin(request, env);
    }
    if (path === '/verify' && request.method === 'POST') {
      return handleVerify(request, env, ctx);
    }
    if (path === '/logout') {
      return handleLogout(env, sessionToken);
    }

    // --- 受保护的路由 (Protected Routes) ---
    if (!session) {
      if (path === '/login-page') {
        return new Response(getLoginPage(env.APP_TITLE), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
      }
      return Response.redirect(new URL('/login-page', request.url).toString(), 302);
    }

    // --- 已登录的路由 (Logged-in Routes) ---

    // 1. 主页 (GET /)
    if (path === '/' && request.method === 'GET') {
      return new Response(getUploadPage(env.APP_TITLE, session.email), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }

    // 2. 创建上传会话 (POST /create-upload-session)
    if (path === '/create-upload-session' && request.method === 'POST') {
      return handleCreateUploadSession(request, env, session);
    }

    // 3. 上传分片 (POST /upload-chunk)
    if (path === '/upload-chunk' && request.method === 'POST') {
      return handleUploadChunk(request, env, session);
    }

    // 4. 完成/取消上传 (POST /complete-upload)
    if (path === '/complete-upload' && request.method === 'POST') {
      return handleCompleteUpload(request, env, ctx, session);
    }
    
    // 5. 成绩汇总下载 (GET /summary)
    if (path === '/summary' && request.method === 'GET') {
      const homeworkName = url.searchParams.get('homework');
      if (!homeworkName) {
        return new Response(JSON.stringify({ error: '缺少 "homework" 参数' }), { status: 400, headers: { 'Content-Type': 'application/json' } });
      }
      return handleSummaryDownload(env, homeworkName);
    }

    // 默认 404
    return new Response('Not Found', { status: 404 });
  },
};

// --- 认证和会话 (Auth and Session) ---
// [!!] handleLogin, handleVerify, handleLogout, isAuthorized, getSession, getCookie
// [!!] sendEmail, getMsGraphToken (这些函数与 v3 版本相同，此处省略以保持清晰)

async function handleLogin(request, env) {
  try {
    const { email } = await request.json();
    if (!email) {
      return new Response(JSON.stringify({ error: 'Email is required.' }), { status: 400, headers: { 'Content-Type': 'application/json' } });
    }
    if (!isAuthorized(email, env.TEACHER_WHITELIST)) {
      return new Response(JSON.stringify({ error: 'Email address or domain is not authorized.' }), { status: 403, headers: { 'Content-Type': 'application/json' } });
    }
    const code = Math.floor(100000 + Math.random() * 900000).toString();
    const codeKey = `code:${email}`;
    await env.AUTH_KV.put(codeKey, code, { expirationTtl: CODE_TTL_SECONDS });
    const mailSent = await sendEmail(
      env.MS_CLIENT_ID,
      env.MS_CLIENT_SECRET,
      env.MS_TENANT_ID,
      env.MS_USER_ID,
      email,
      `[${env.APP_TITLE}] 您的登录验证码`,
      `您的登录验证码是：<b>${code}</b><br>此验证码将在5分钟内失效。`
    );
    if (!mailSent) {
      return new Response(JSON.stringify({ error: 'Failed to send email.' }), { status: 500, headers: { 'Content-Type': 'application/json' } });
    }
    return new Response(JSON.stringify({ success: true, message: `Verification code sent to ${email}.` }), { status: 200, headers: { 'Content-Type': 'application/json' } });
  } catch (err) {
    console.error(`Login error: ${err.message}`);
    return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { 'Content-Type': 'application/json' } });
  }
}

async function handleVerify(request, env, ctx) {
  try {
    const { email, code } = await request.json();
    if (!email || !code) {
      return new Response(JSON.stringify({ error: 'Email and code are required.' }), { status: 400, headers: { 'Content-Type': 'application/json' } });
    }
    const codeKey = `code:${email}`;
    const storedCode = await env.AUTH_KV.get(codeKey);
    if (!storedCode || storedCode !== code) {
      return new Response(JSON.stringify({ error: 'Invalid or expired code.' }), { status: 403, headers: { 'Content-Type': 'application/json' } });
    }
    const sessionToken = crypto.randomUUID();
    const sessionKey = `session:${sessionToken}`;
    const sessionData = { email: email, createdAt: Date.now() };
    await env.AUTH_KV.put(sessionKey, JSON.stringify(sessionData), { expirationTtl: SESSION_DURATION_SECONDS });
    ctx.waitUntil(env.AUTH_KV.delete(codeKey));
    const cookie = `${SESSION_COOKIE_NAME}=${sessionToken}; HttpOnly; Secure; Path=/; Max-Age=${SESSION_DURATION_SECONDS}; SameSite=Strict`;
    return new Response(JSON.stringify({ success: true }), {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Set-Cookie': cookie },
    });
  } catch (err) {
    console.error(`Verify error: ${err.message}`);
    return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { 'Content-Type': 'application/json' } });
  }
}

async function handleLogout(env, sessionToken) {
  if (sessionToken) {
    await env.AUTH_KV.delete(`session:${sessionToken}`);
  }
  const cookie = `${SESSION_COOKIE_NAME}=; HttpOnly; Secure; Path=/; Max-Age=0; SameSite=Strict`;
  return new Response(null, {
    status: 302,
    headers: { 'Set-Cookie': cookie, 'Location': '/login-page' },
  });
}

function isAuthorized(email, whitelist) {
  if (!whitelist) return false;
  const emailLower = email.toLowerCase();
  const rules = whitelist.split(',')
    .map(rule => rule.trim().toLowerCase()) 
    .filter(rule => rule.length > 0);
  for (const rule of rules) {
    if (rule.startsWith('@')) {
      if (emailLower.endsWith(rule)) return true;
    } else {
      if (emailLower === rule) return true;
    }
  }
  return false;
}

async function getSession(kv, sessionToken) {
  if (!sessionToken) return null;
  const sessionKey = `session:${sessionToken}`;
  const data = await kv.get(sessionKey);
  if (!data) return null;
  return JSON.parse(data);
}

function getCookie(cookieString, name) {
  const match = cookieString.match(new RegExp('(^| )' + name + '=([^;]+)'));
  return match ? match[2] : null;
}

async function sendEmail(clientId, clientSecret, tenantId, fromUser, toUser, subject, content) {
  try {
    const token = await getMsGraphToken(clientId, clientSecret, tenantId);
    if (!token) throw new Error('Failed to get MS Graph token for sending mail.');
    const sendMailUrl = `https://graph.microsoft.com/v1.0/users/${fromUser}/sendMail`;
    const emailBody = {
      message: {
        subject: subject,
        body: { contentType: 'HTML', content: content },
        toRecipients: [{ emailAddress: { address: toUser } }],
      },
      saveToSentItems: 'false',
    };
    const response = await fetch(sendMailUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(emailBody),
    });
    return response.ok;
  } catch (err) {
    console.error(`Email send failed: ${err.message}`);
    return false;
  }
}

// --- 核心业务逻辑 (Core Business Logic) V4 ---

// 1. 创建上传会话
async function handleCreateUploadSession(request, env, session) {
  try {
    const { fileName } = await request.json();
    if (!fileName) {
      return new Response(JSON.stringify({ error: 'fileName is required.' }), { status: 400 });
    }

    const token = await getMsGraphToken(env.MS_CLIENT_ID, env.MS_CLIENT_SECRET, env.MS_TENANT_ID);
    if (!token) throw new Error('无法获取 MS Graph token。');

    const homeworkName = fileName.endsWith('.zip') ? fileName.slice(0, -4) : fileName;
    const odPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}`; // 这是作业文件夹
    const odZipPath = `${odPath}/${fileName}`; // 这是 zip 文件本身

    // [V4] 创建一个可续传上传会话
    const createSessionUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${odZipPath}:/createUploadSession`;
    
    const sessionResponse = await fetch(createSessionUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        "item": {
          "@microsoft.graph.conflictBehavior": "replace"
        }
      })
    });

    if (!sessionResponse.ok) {
      throw new Error(`OneDrive create session failed: ${await sessionResponse.text()}`);
    }

    const { uploadUrl } = await sessionResponse.json();
    
    // [!!] 将上传 URL 存储在 KV 中，以便分片时使用
    const sessionId = crypto.randomUUID();
    const sessionKey = `upload:${session.email}:${sessionId}`;
    await env.AUTH_KV.put(sessionKey, JSON.stringify({ uploadUrl, homeworkName, odPath, odZipPath }), { expirationTtl: UPLOAD_SESSION_TTL_SECONDS });
    
    console.log(`[Upload Session] Created for ${homeworkName}`);
    return new Response(JSON.stringify({ success: true, sessionId }), { status: 200 });

  } catch (err) {
    console.error(`[Create Session Error] ${err.message}`);
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
}

// 2. 代理上传分片
async function handleUploadChunk(request, env, session) {
  try {
    const sessionId = request.headers.get('X-Session-Id');
    const contentRange = request.headers.get('Content-Range'); // e.g., "bytes 0-999/10000"
    const contentLength = request.headers.get('Content-Length');

    if (!sessionId || !contentRange || !contentLength) {
      return new Response(JSON.stringify({ error: 'Missing headers: X-Session-Id, Content-Range, Content-Length' }), { status: 400 });
    }
    
    const sessionKey = `upload:${session.email}:${sessionId}`;
    const sessionData = await env.AUTH_KV.get(sessionKey, 'json');

    if (!sessionData) {
      return new Response(JSON.stringify({ error: 'Invalid or expired upload session.' }), { status: 404 });
    }
    
    // [!!] 代理请求: 将分片流式传输到 OneDrive
    const { uploadUrl } = sessionData;
    
    const uploadResponse = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Range': contentRange,
        'Content-Length': contentLength,
      },
      body: request.body, // [!!] 流式传输 (Stream)
    });

    if (!uploadResponse.ok) {
      // 检查是否是 200, 201, 202 (Accepted) 或 204 (No Content)
      // 如果上传完成，它会返回 201 Created 或 200 OK
      if (uploadResponse.status === 201 || uploadResponse.status === 200) {
         // 这是最后一个分片，上传已完成
         const data = await uploadResponse.json();
         // 最后一个分片已成功，但我们让前端来触发 /complete-upload
         return new Response(JSON.stringify({ success: true, ...data }), { status: 200 });
      }
      throw new Error(`OneDrive chunk upload failed: ${uploadResponse.status} ${await uploadResponse.text()}`);
    }

    // 202 Accepted (分片已接收，但未完成)
    const data = await uploadResponse.json();
    return new Response(JSON.stringify({ success: true, ...data }), { status: 202 });

  } catch (err) {
    console.error(`[Upload Chunk Error] ${err.message}`);
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
}

// 3. 完成上传（或取消）并触发后台任务
async function handleCompleteUpload(request, env, ctx, session) {
  const { sessionId, status } = await request.json();
  const sessionKey = `upload:${session.email}:${sessionId}`;
  const sessionData = await env.AUTH_KV.get(sessionKey, 'json');

  if (!sessionData) {
    return new Response(JSON.stringify({ error: 'Invalid or expired upload session.' }), { status: 404 });
  }

  // 无论成功还是失败，都删除 KV 中的会话
  ctx.waitUntil(env.AUTH_KV.delete(sessionKey));

  if (status === 'cancelled') {
    // 如果取消，则异步删除 OneDrive 上的上传会话
    ctx.waitUntil(fetch(sessionData.uploadUrl, { method: 'DELETE' }));
    console.log(`[Upload Session] Cancelled by user: ${sessionData.homeworkName}`);
    return new Response(JSON.stringify({ success: true, message: 'Upload cancelled.' }), { status: 200 });
  }

  if (status === 'completed') {
    // [!!] 异步执行 (Execute Asynchronously)
    // 立即返回响应，防止浏览器超时
    ctx.waitUntil(processGradingInBackgroundV4(sessionData, env));
    
    console.log(`[Upload Session] Completed. Starting background job: ${sessionData.homeworkName}`);
    return new Response(JSON.stringify({ success: true, message: `文件 "${sessionData.homeworkName}.zip" 已收到，正在后台处理批改。` }), { status: 202 });
  }
  
  return new Response(JSON.stringify({ error: 'Invalid status.' }), { status: 400 });
}


// [V4] 后台处理 (使用远程解压)
async function processGradingInBackgroundV4(sessionData, env) {
  const { homeworkName, odPath, odZipPath } = sessionData;
  console.log(`[V4 开始处理] ${homeworkName}`);
  
  try {
    // 1. 获取 MS Graph API Token
    const token = await getMsGraphToken(env.MS_CLIENT_ID, env.MS_CLIENT_SECRET, env.MS_TENANT_ID);
    if (!token) throw new Error('无法获取 MS Graph token。');

    // 2. [!!] 命令 OneDrive 远程解压 ZIP
    // a. 获取 Zip 文件的 DriveItem ID
    const zipItem = await getDriveItem(token, odZipPath);
    if (!zipItem || !zipItem.id) throw new Error(`无法在 OneDrive 上找到 Zip 文件: ${odZipPath}`);

    // b. 创建解压目标文件夹 (如果 /extract 不自动创建的话)
    // (通常 /extract 会自动创建以 zip 文件名命名的文件夹，但我们先确保父文件夹存在)
    await createOneDriveFolder(token, odPath); // 确保作业文件夹存在
    
    console.log(`[V4 远程解压] 开始解压 ${homeworkName}.zip`);

    // c. 调用 /extract API
    // 这会将 student1.zip, student2.zip... 解压到 odPath (即 '.../HomeworkGrader/MyHomework/')
    const extractUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${zipItem.id}/extract`;
    const extractResponse = await fetch(extractUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { path: `/drive/root:/${odPath}` } }) // 解压到父文件夹
    });

    if (!extractResponse.ok && extractResponse.status !== 202) {
      throw new Error(`OneDrive /extract failed: ${await extractResponse.text()}`);
    }
    
    // (注意: /extract 是一个长轮询操作，但对于 worker 来说，我们可能需要轮询)
    // (为简单起见，我们假设它足够快，或者在下一步的 listChildren 中重试)
    // (在实际生产中，这里需要一个轮询逻辑来检查 'location' header 的状态)
    
    // 为简单起见，我们先等待 10 秒
    await new Promise(resolve => setTimeout(resolve, 10000)); // 10s delay for extraction
    
    console.log(`[V4 远程解压] 解压完成 (假设)。开始遍历学生...`);

    // 3. 列出解压后的学生 ZIP 包
    // (现在 odPath 文件夹下应该有 student1.zip, student2.zip...)
    const studentZips = await listDriveChildren(token, odPath);
    if (!studentZips || studentZips.length === 0) {
      throw new Error(`在远程解压目录中未找到学生 .zip 文件: ${odPath}`);
    }

    // 4. 循环处理每个学生 (现在是远程文件)
    for (const studentZipItem of studentZips) {
      if (!studentZipItem.name.endsWith('.zip') || studentZipItem.file === undefined) continue;

      const studentName = studentZipItem.name.slice(0, -4);
      const studentZipId = studentZipItem.id;
      const studentFolderPath = `${odPath}/${studentName}`; // e.g., .../MyHomework/student1

      console.log(`[V4 处理学生] ${studentName}`);

      try {
        // 5. [!!] 再次远程解压学生的 ZIP 包
        await createOneDriveFolder(token, studentFolderPath); // 确保学生文件夹存在
        
        const studentExtractUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${studentZipId}/extract`;
        const studentExtractRes = await fetch(studentExtractUrl, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ parentReference: { path: `/drive/root:/${studentFolderPath}` } })
        });
        
        if (!studentExtractRes.ok && studentExtractRes.status !== 202) {
             throw new Error(`无法解压学生 Zip: ${studentName}: ${await studentExtractRes.text()}`);
        }
        
        // 再次等待解压
        await new Promise(resolve => setTimeout(resolve, 5000)); // 5s delay

        // 6. 列出学生作业文件 (e.g., code.py, report.pdf)
        const studentFilesList = await listDriveChildren(token, studentFolderPath);
        if (!studentFilesList || studentFilesList.length === 0) {
            throw new Error(`学生 ${studentName} 的 .zip 包是空的。`);
        }

        const studentHomeworkFiles = [];
        for (const fileItem of studentFilesList) {
          if (fileItem.folder || fileItem.name.startsWith('.')) continue; // 跳过文件夹

          // 7. [!!] 从 OneDrive 下载单个文件内容 (这很小)
          const fileContent = await getFileContentFromOneDrive(token, `${studentFolderPath}/${fileItem.name}`, false);
          if (fileContent) {
            studentHomeworkFiles.push({
              name: fileItem.name,
              type: fileItem.file.mimeType || 'application/octet-stream',
              content: fileContent, // 这是 Blob
            });
          }
        }

        // 8. 调用 Gemini API 进行批改
        const reportJson = await gradeHomeworkWithGemini(studentHomeworkFiles, homeworkName, studentName, env.GOOGLE_API_KEY);
        
        // 9. 上传报告到 OneDrive
        if (reportJson) {
          const reportPath = `${studentFolderPath}/report.json`;
          await uploadFileToOneDrive(token, reportPath, JSON.stringify(reportJson, null, 2), 'application/json');
          console.log(`[V4 报告上传成功] ${studentName}`);
        } else {
          console.warn(`[V4 Gemini 未返回报告] ${studentName}`);
        }
      } catch (studentErr) {
        console.error(`[V4 处理学生失败] ${studentName}: ${studentErr.message}`);
        const errorReportPath = `${odPath}/${studentName}/error.json`;
        await uploadFileToOneDrive(token, errorReportPath, JSON.stringify({ error: studentErr.message }), 'application/json');
      }
    }
    
    // (可选) 删除上传的主 Zip 文件
    ctx.waitUntil(deleteDriveItem(token, zipItem.id));
    console.log(`[V4 处理完成] ${homeworkName}`);

  } catch (err) {
    console.error(`[V4 后台任务失败] ${homeworkName}: ${err.message}`);
  }
}

// [V4] 汇总下载 (逻辑与 V3 相同)
async function handleSummaryDownload(env, homeworkName) {
  try {
    console.log(`[下载汇总] 开始: ${homeworkName}`);
    const token = await getMsGraphToken(env.MS_CLIENT_ID, env.MS_CLIENT_SECRET, env.MS_TENANT_ID);
    if (!token) throw new Error('无法获取 MS Graph token。');

    const homeworkPath = `${env.MS_ONEDRIVE_BASE_PATH}/${homeworkName}`;

    const studentFolders = await listDriveChildren(token, homeworkPath);
    if (!studentFolders || studentFolders.length === 0) {
      throw new Error('未找到学生提交记录。请确保 AI 批改已完成。');
    }
    
    const csvData = [];
    csvData.push(['StudentName', 'OverallGrade', 'OverallFeedback']); // CSV Header

    for (const folder of studentFolders) {
      // [V4] 确保我们只处理文件夹，并跳过 .zip 文件
      if (!folder.folder || folder.name.endsWith('.zip')) continue; 
      
      const studentName = folder.name;
      const reportPath = `${homeworkPath}/${studentName}/report.json`;
      const reportContent = await getFileContentFromOneDrive(token, reportPath, true); // true = asText

      if (reportContent) {
        try {
          const report = JSON.parse(reportContent);
          csvData.push([
            `"${studentName}"`,
            report.overall_grade || 'N/A',
            `"${(report.overall_feedback || 'N/A').replace(/"/g, '""')}"`
          ]);
        } catch (e) {
          csvData.push([`"${studentName}"`, '报告解析失败', 'N/A']);
        }
      } else {
        const errorPath = `${homeworkPath}/${studentName}/error.json`;
        const errorContent = await getFileContentFromOneDrive(token, errorPath, true);
        if (errorContent) {
           try {
             const errorData = JSON.parse(errorContent);
             csvData.push([`"${studentName}"`, '批改出错', `"${(errorData.error || errorContent).replace(/"/g, '""')}"`]);
           } catch(e) {
             csvData.push([`"${studentName}"`, '批改出错', `"${errorContent.replace(/"/g, '""')}"`]);
           }
        } else {
          csvData.push([`"${studentName}"`, '批改未完成', 'N/A']);
        }
      }
    }

    const csvString = csvData.map(row => row.join(',')).join('\r\n');
    const csvFileName = `summary_${homeworkName.replace(/[^a-z0-9]/gi, '_')}.csv`;
    
    console.log(`[下载汇总] 成功: ${homeworkName}`);
    return new Response(csvString, {
      headers: {
        'Content-Type': 'text/csv; charset=utf-8',
        'Content-Disposition': `attachment; filename="${csvFileName}"`
      }
    });

  } catch (err) {
    console.error(`[下载汇总失败] ${err.message}`);
    return new Response(JSON.stringify({ error: `下载汇总失败: ${err.message}` }), { status: 500, headers: { 'Content-Type': 'application/json' } });
  }
}


// --- Gemini API (AI 批改) V4 ---
// (与 V3 基本相同，但 content 是 Blob)
async function gradeHomeworkWithGemini(files, homeworkName, studentName, apiKey) {
  try {
    console.log(`[Gemini 请求] 正在批改 ${studentName}`);
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

    const generationConfig = { responseMimeType: 'application/json', temperature: 0.2 };
    const safetySettings = [
      { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
    ];
    
    const prompt_parts = [
      { "text": `
# 角色
你是一位严格、细致的AI助教。
# 任务
批改一份来自学生 "${studentName}" 的作业，作业主题是 "${homeworkName}"。
学生的提交材料在下面提供，可能是代码、文档、图片、音视频等。
# 批改要求
1.  **综合分析**：你必须查看学生提交的 **所有** 文件，综合评估。
2.  **给出分数**：给出一个 0-100 之间的总分 (overall_grade)。
3.  **给出评语**：给出一个简短、有针对性的总体评语 (overall_feedback)，指出优点和主要问题。
# 输出格式
严格按照以下 JSON 格式输出，不要包含任何 markdown 标记。
{
  "student_name": "${studentName}",
  "overall_grade": 85,
  "overall_feedback": "做得不错。代码逻辑基本正确，但缺少对边缘情况的处理。项目报告的分析比较到位。"
}
` }
    ];

    for (const file of files) {
      if (file.content.size > 0) {
        prompt_parts.push({
          inlineData: {
            mimeType: file.type,
            data: Buffer.from(await file.content.arrayBuffer()).toString('base64'),
          },
        });
      }
    }
    
    const result = await model.generateContent({
      contents: [{ role: 'user', parts: prompt_parts }],
      generationConfig,
      safetySettings,
    });

    const response = result.response;
    const jsonText = response.text();
    return JSON.parse(jsonText); 

  } catch (err) {
    console.error(`[Gemini 失败] ${studentName}: ${err.message}`);
    if (err.response && err.response.promptFeedback) {
      console.error('[Gemini 安全阻挡]', JSON.stringify(err.response.promptFeedback));
    }
    return null;
  }
}


// --- Microsoft Graph API 助手 (Helpers) V4 ---

async function getMsGraphToken(clientId, clientSecret, tenantId) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });
  try {
    const response = await fetch(url, {
      method: 'POST',
      body: body,
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Token request failed: ${response.status} ${response.statusText} - ${errorText}`);
    }
    const data = await response.json();
    console.log('[MS Token 成功] (应用流)');
    return data.access_token;
  } catch (err) {
    console.error(`[MS Token 失败] ${err.message}`);
    return null;
  }
}

async function uploadFileToOneDrive(token, pathOnOneDrive, content, contentType) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/content`;
  const response = await fetch(url, {
    method: 'PUT',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': contentType },
    body: content,
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`OneDrive upload failed: ${response.status} ${response.statusText} - ${errorText}`);
  }
  return await response.json();
}

async function getFileContentFromOneDrive(token, pathOnOneDrive, asText = true) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/content`;
  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`Status ${response.status}`);
    return asText ? await response.text() : await response.blob();
  } catch (err) {
    console.warn(`[OneDrive 下载失败] ${pathOnOneDrive}: ${err.message}`);
    return null;
  }
}

async function getDriveItem(token, pathOnOneDrive) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}`;
  try {
    const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`GetItem Status ${response.status}`);
    return await response.json();
  } catch (err) {
    console.error(`[GetDriveItem 失败] ${pathOnOneDrive}: ${err.message}`);
    return null;
  }
}

async function listDriveChildren(token, pathOnOneDrive) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/children?$select=name,folder,file,id`;
  try {
    const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`ListChildren Status ${response.status}`);
    const { value } = await response.json();
    return value || [];
  } catch (err) {
    console.error(`[ListChildren 失败] ${pathOnOneDrive}: ${err.message}`);
    return [];
  }
}

async function createOneDriveFolder(token, pathOnOneDrive) {
    const parts = pathOnOneDrive.split('/');
    let currentPath = '';
    // 我们假设 MS_ONEDRIVE_BASE_PATH (e.g., 'Apps/HomeworkGrader') 已经存在
    // 我们只需要创建 '.../HomeworkGrader/HomeworkName' 和 '.../HomeworkGrader/HomeworkName/StudentName'
    const basePathParts = env.MS_ONEDRIVE_BASE_PATH.split('/').length;
    
    // 我们从 base path 之后开始创建
    let pathToCheck = env.MS_ONEDRIVE_BASE_PATH;
    
    for (let i = basePathParts; i < parts.length; i++) {
        pathToCheck = `${pathToCheck}/${parts[i]}`;
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathToCheck}`;
        // 尝试获取
        const getRes = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
        
        if (getRes.status === 404) {
          // 不存在，创建它
          const parentPath = parts.slice(0, i).join('/');
          const createUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${parentPath}:/children`;
          const createRes = await fetch(createUrl, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ "name": parts[i], "folder": {}, "@microsoft.graph.conflictBehavior": "fail" })
          });
          if (!createRes.ok) {
             // 忽略 409 Conflict (已存在)
             if (createRes.status !== 409) {
                 console.warn(`[CreateFolder] Failed to create ${parts[i]}: ${await createRes.text()}`);
             }
          }
        }
    }
}

async function deleteDriveItem(token, itemId) {
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`;
    try {
        const response = await fetch(url, { method: 'DELETE', headers: { 'Authorization': `Bearer ${token}` } });
        if (response.ok) console.log(`[DeleteDriveItem] 成功删除 ${itemId}`);
    } catch (err) {
        console.error(`[DeleteDriveItem 失败] ${itemId}: ${err.message}`);
    }
}


// --- HTML 页面 (HTML Pages) V4 ---

function getLoginPage(appTitle) {
  // [!!] V4 登录页 (与 V3 相同)
  return `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${appTitle} - 登录</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 flex items-center justify-center min-h-screen">
    <div class="bg-white p-8 rounded-lg shadow-md w-full max-w-sm">
        <h1 class="text-2xl font-bold text-center mb-6">${appTitle}</h1>
        <form id="loginForm" class="space-y-4">
            <div>
                <label for="email" class="block text-sm font-medium text-gray-700">教师邮箱</label>
                <input type="email" id="email" name="email" required class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500">
            </div>
            <button type="submit" id="sendCodeBtn" class="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500">
                发送验证码
            </button>
        </form>
        <form id="verifyForm" class="space-y-4 hidden">
            <div class="text-sm text-gray-600">已发送验证码至 <b id="emailDisplay"></b></div>
            <div>
                <label for="code" class="block text-sm font-medium text-gray-700">验证码</label>
                <input type="text" id="code" name="code" required inputmode="numeric" pattern="[0-9]{6}" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500">
            </div>
            <button type="submit" id="verifyBtn" class="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500">
                登录
            </button>
            <button type="button" id="backBtn" class="w-full text-center text-sm text-gray-500 hover:text-gray-700">返回</button>
        </form>
        <div id="message" class="mt-4 text-center text-sm"></div>
    </div>
    <script>
        const loginForm = document.getElementById('loginForm');
        const verifyForm = document.getElementById('verifyForm');
        const sendCodeBtn = document.getElementById('sendCodeBtn');
        const verifyBtn = document.getElementById('verifyBtn');
        const backBtn = document.getElementById('backBtn');
        const emailInput = document.getElementById('email');
        const emailDisplay = document.getElementById('emailDisplay');
        const codeInput = document.getElementById('code');
        const message = document.getElementById('message');

        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const email = emailInput.value;
            sendCodeBtn.disabled = true; sendCodeBtn.textContent = '发送中...'; message.textContent = '';
            try {
                const res = await fetch('/login', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email })
                });
                if (!res.ok) { const { error } = await res.json(); throw new Error(error || '发送失败'); }
                message.textContent = '验证码已发送，请查收邮件。';
                message.className = 'mt-4 text-center text-sm text-green-600';
                emailDisplay.textContent = email;
                loginForm.classList.add('hidden');
                verifyForm.classList.remove('hidden');
            } catch (err) {
                message.textContent = '错误: ' + err.message;
                message.className = 'mt-4 text-center text-sm text-red-600';
            } finally {
                sendCodeBtn.disabled = false; sendCodeBtn.textContent = '发送验证码';
            }
        });
        verifyForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const email = emailInput.value; const code = codeInput.value;
            verifyBtn.disabled = true; verifyBtn.textContent = '登录中...'; message.textContent = '';
            try {
                const res = await fetch('/verify', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email, code })
                });
                if (!res.ok) { const { error } = await res.json(); throw new Error(error || '登录失败'); }
                message.textContent = '登录成功！正在跳转...';
                message.className = 'mt-4 text-center text-sm text-green-600';
                window.location.href = '/';
            } catch (err) {
                message.textContent = '错误: ' + err.message;
                message.className = 'mt-4 text-center text-sm text-red-600';
            } finally {
                verifyBtn.disabled = false; verifyBtn.textContent = '登录';
            }
        });
        backBtn.addEventListener('click', () => {
            loginForm.classList.remove('hidden'); verifyForm.classList.add('hidden'); message.textContent = '';
        });
    </script>
</body>
</html>
      `;
}

function getUploadPage(appTitle, userEmail) {
  // [!!] V4 主页 (实现了分片上传)
  return `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${appTitle} - 主页</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 min-h-screen p-4 md:p-8">
    <div class="max-w-3xl mx-auto">
        
        <!-- Header -->
        <div class="flex justify-between items-center mb-6">
            <h1 class="text-3xl font-bold text-gray-800">${appTitle}</h1>
            <div class="text-right">
                <span class="text-sm text-gray-600">${userEmail}</span>
                <a href="/logout" class="ml-4 text-sm text-blue-600 hover:underline">退出登录</a>
            </div>
        </div>

        <!-- Status Message -->
        <div id="message" class.bind="hidden p-4 rounded-md mb-6"></div>

        <!-- 1. Upload Section -->
        <div class="bg-white p-6 rounded-lg shadow-md mb-8">
            <h2 class="text-2xl font-semibold mb-4">1. 上传作业 ZIP 包 (支持大文件)</h2>
            <form id="uploadForm" class="space-y-4">
                <div>
                    <label for="homeworkFile" class="block text-sm font-medium text-gray-700">选择主 .zip 文件</label>
                    <p class="text-xs text-gray-500 mb-2">(.zip 包内应包含所有学生的 .zip 文件。支持 > 100MB 文件)</p>
                    <input type="file" id="homeworkFile" name="homeworkFile" required accept=".zip"
                        class="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-gray-50
                               file:mr-4 file:py-2 file:px-4 file:rounded-l-lg file:border-0 file:text-sm file:font-semibold
                               file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100">
                </div>
                
                <!-- Progress Bar -->
                <div id="progressContainer" class="w-full bg-gray-200 rounded-full h-2.5 hidden">
                    <div id="progressBar" class="bg-blue-600 h-2.5 rounded-full" style="width: 0%"></div>
                </div>
                
                <div class="flex space-x-4">
                    <button type="submit" id="uploadBtn"
                        class="w-1/2 flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:opacity-50">
                        上传并批改
                    </button>
                    <button type="button" id="cancelBtn"
                        class="w-1/2 flex justify-center py-2 px-4 border border-gray-300 rounded-md shadow-sm text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-gray-500 disabled:opacity-50"
                        disabled>
                        取消上传
                    </button>
                </div>
            </form>
        </div>

        <!-- 2. Download Section -->
        <div class="bg-white p-6 rounded-lg shadow-md">
            <h2 class="text-2xl font-semibold mb-4">2. 下载成绩总表 (CSV)</h2>
            <form id="downloadForm" class="space-y-4">
                <div>
                    <label for="homeworkName" class="block text-sm font-medium text-gray-700">作业名称</label>
                    <p class="text-xs text-gray-500 mb-2">(即您上传的 .zip 文件名，不含 .zip 后缀)</p>
                    <input type="text" id="homeworkName" name="homeworkName" required
                        class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500">
                </div>
                <button type="submit" id="downloadBtn"
                    class="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500 disabled:opacity-50">
                    下载汇总
                </button>
            </form>
        </div>
    </div>

    <script>
        const uploadForm = document.getElementById('uploadForm');
        const uploadBtn = document.getElementById('uploadBtn');
        const cancelBtn = document.getElementById('cancelBtn');
        const downloadForm = document.getElementById('downloadForm');
        const downloadBtn = document.getElementById('downloadBtn');
        const message = document.getElementById('message');
        const homeworkNameInput = document.getElementById('homeworkName');
        const fileInput = document.getElementById('homeworkFile');
        const progressContainer = document.getElementById('progressContainer');
        const progressBar = document.getElementById('progressBar');

        // [V4] 分片上传全局变量
        let currentUpload = {
            sessionId: null,
            file: null,
            isCancelled: false,
            chunkSize: 10 * 1024 * 1024, // 10MB Chunks
        };

        // --- Upload Handler (V4) ---
        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            if (!fileInput.files || fileInput.files.length === 0) {
                showMessage('请先选择一个 .zip 文件。', 'red');
                return;
            }
            
            const file = fileInput.files[0];
            const fileName = file.name;
            currentUpload.file = file;
            currentUpload.isCancelled = false;

            // 自动填充下载框
            if (fileName.toLowerCase().endsWith('.zip')) {
                homeworkNameInput.value = fileName.slice(0, -4);
            }

            uploadBtn.disabled = true;
            cancelBtn.disabled = false;
            uploadBtn.textContent = '上传中... (0%)';
            showMessage('正在创建上传会话...', 'blue');
            progressContainer.classList.remove('hidden');
            progressBar.style.width = '0%';

            try {
                // 1. 创建上传会话
                const createRes = await fetch('/create-upload-session', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ fileName: fileName })
                });

                if (!createRes.ok) throw new Error('无法创建上传会话: ' + (await createRes.json()).error);
                
                const { sessionId } = await createRes.json();
                currentUpload.sessionId = sessionId;

                // 2. 开始分片上传
                await uploadChunks();
                
                // 3. 完成上传
                if (!currentUpload.isCancelled) {
                    await completeUpload(sessionId, 'completed');
                }

            } catch (err) {
                if (!currentUpload.isCancelled) {
                    showMessage('上传失败: ' + err.message, 'red');
                }
                if (currentUpload.sessionId) {
                    await completeUpload(currentUpload.sessionId, 'cancelled');
                }
            } finally {
                resetUploadState();
            }
        });

        async function uploadChunks() {
            const { file, sessionId, chunkSize } = currentUpload;
            const totalChunks = Math.ceil(file.size / chunkSize);

            for (let chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++) {
                if (currentUpload.isCancelled) {
                    throw new Error('Upload cancelled by user.');
                }

                const start = chunkIndex * chunkSize;
                const end = Math.min(start + chunkSize, file.size);
                const chunk = file.slice(start, end);
                const chunkNum = chunkIndex + 1;

                const progress = Math.round((start / file.size) * 100);
                uploadBtn.textContent = \`上传中... (\${progress}%)\`;
                progressBar.style.width = \`\${progress}%\`;
                showMessage(\`正在上传分片 \${chunkNum} / \${totalChunks}...\`, 'blue');

                const contentRange = \`bytes \${start}-\${end - 1}/\${file.size}\`;

                const res = await fetch('/upload-chunk', {
                    method: 'POST',
                    headers: {
                        'X-Session-Id': sessionId,
                        'Content-Range': contentRange,
                        'Content-Length': chunk.size,
                    },
                    body: chunk
                });

                if (!res.ok) {
                    // 200/201 意味着最后的分片已完成
                    if (res.status === 200 || res.status === 201) {
                         break; // 上传完成
                    }
                    // 202 意味着分片成功，但未完成
                    if (res.status !== 202) {
                        throw new Error(\`分片 \${chunkNum} 上传失败: \${(await res.json()).error}\`);
                    }
                }
                // 否则 (202), 继续循环
            }
            
            // 确保进度条达到 100%
            uploadBtn.textContent = '上传中... (100%)';
            progressBar.style.width = '100%';
        }

        async function completeUpload(sessionId, status) {
            showMessage('文件已传完，正在通知服务器处理...', 'blue');
            try {
                const res = await fetch('/complete-upload', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sessionId, status })
                });
                
                const data = await res.json();
                if (!res.ok) throw new Error(data.error || '完成上传时出错');

                if (status === 'completed') {
                    showMessage(data.message, 'green'); // "文件已收到，正在后台处理..."
                    uploadForm.reset();
                } else {
                    showMessage('上传已取消。', 'blue');
                }
            } catch (err) {
                 if (status !== 'cancelled') {
                    showMessage(\`完成上传失败: \${err.message}\`, 'red');
                 }
            }
        }
        
        cancelBtn.addEventListener('click', () => {
            currentUpload.isCancelled = true;
            resetUploadState();
            showMessage('上传已取消。', 'blue');
        });
        
        function resetUploadState() {
            uploadBtn.disabled = false;
            cancelBtn.disabled = true;
            uploadBtn.textContent = '上传并批改';
            progressContainer.classList.add('hidden');
            progressBar.style.width = '0%';
            currentUpload.sessionId = null;
            currentUpload.file = null;
            currentUpload.isCancelled = false;
        }

        // --- Download Handler (V4) ---
        // (此部分与 V3 相同)
        downloadForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const homeworkName = homeworkNameInput.value;
            if (!homeworkName) {
                showMessage('请输入作业名称。', 'red');
                return;
            }
            downloadBtn.disabled = true; downloadBtn.textContent = '正在生成...';
            showMessage('正在抓取和汇总报告...', 'blue');
            try {
                const res = await fetch(\`/summary?homework=\${encodeURIComponent(homeworkName)}\`);
                if (!res.ok) {
                    const data = await res.json();
                    throw new Error(data.error || '下载失败');
                }
                const blob = await res.blob();
                const header = res.headers.get('Content-Disposition');
                const filenameMatch = header && header.match(/filename="(.+?)"/);
                const filename = filenameMatch ? filenameMatch[1] : \`summary_\${homeworkName}.csv\`;
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = filename;
                document.body.appendChild(a); a.click();
                a.remove(); window.URL.revokeObjectURL(url);
                showMessage('CSV 汇总表已开始下载。', 'green');
            } catch (err) {
                showMessage('下载失败: ' + err.message, 'red');
            } finally {
                downloadBtn.disabled = false; downloadBtn.textContent = '下载汇总';
            }
        });

        function showMessage(text, color) {
            message.textContent = text;
            message.className = \`p-4 rounded-md mb-6 text-white \${color === 'red' ? 'bg-red-500' : (color === 'green' ? 'bg-green-500' : 'bg-blue-500')}\`;
        }
    </script>
</body>
</html>
      `;
}

