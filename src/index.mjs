/**
 * Cloudflare Worker - AI 自动作业批改平台 (v4 - 分片代理 + 远程解压)
 *
 * 功能:
 * 1. 教师邮箱验证码登录 (KV)。
 * 2. [!!] 支持 >100MB 大型文件上传:
 * - (前端) 浏览器将文件分片 (10MB/片)。
 * - (前端) POST /api/create-upload-session: 向 OneDrive 请求上传 URL。
 * - (前端) POST /api/upload-chunk: 浏览器逐片上传。
 * - (后端) Worker 作为代理, 将分片流式传输到 OneDrive URL。
 * 3. 异步后台处理作业 (无内存限制):
 * - (后台) Worker 命令 OneDrive 远程解压 (Graph API /extract)。
 * - (后台) Worker 远程遍历文件夹, 逐个下载小文件。
 * - (后台) 调用 Gemini API 批改。
 * - (后台) 上传 report.json 回 OneDrive。
 * 4. 教师下载 CSV 成绩总表。
 *
 * 依赖:
 * - KV (AUTH_KV): 存储登录验证码和会话。
 * - Secrets: (如 deploy.yml 中所列)
 */

// [!!] 移除了 unzipit
import { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } from '@google/generative-ai';

// -----------------------------------------------------------------
// 1. 主入口: 路由处理
// -----------------------------------------------------------------

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const path = url.pathname;

    // 尝试获取会话 Cookie
    const sessionCookie = getCookie(request, 'auth-session');
    const isAuthenticated = sessionCookie ? await isSessionValid(env.AUTH_KV, sessionCookie) : false;
    const userEmail = isAuthenticated ? await env.AUTH_KV.get(`session:${sessionCookie}`) : null;

    // --- 公共路由 ---
    if (path === '/login') {
      if (request.method === 'POST') return handleLogin(request, env);
      return new Response(getLoginPage(env.APP_TITLE || "AI 批改平台"), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }
    if (path === '/verify') {
      if (request.method === 'POST') return handleVerify(request, env);
      return new Response('Invalid method', { status: 405 });
    }
    if (path === '/logout') {
      return handleLogout(request);
    }

    // --- 以下路由需要认证 ---
    if (!isAuthenticated) {
      // API 请求返回 401
      if (path.startsWith('/api/')) {
        return new Response(JSON.stringify({ success: false, message: 'Unauthorized' }), { status: 401, headers: { 'Content-Type': 'application/json' } });
      }
      // 页面请求重定向到登录
      return Response.redirect(new URL('/login', request.url).toString(), 302);
    }

    // 路由 1: 根路径 (GET - 主页)
    if (path === '/') {
      if (request.method === 'GET') {
        return new Response(getHtmlPage(env.APP_TITLE || "AI 批改平台"), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
      }
      return new Response('Invalid method', { status: 405 });
    }

    // 路由 2: 成绩汇总下载
    if (path === '/summary') {
      if (request.method === 'GET') {
        const homeworkName = url.searchParams.get('homework');
        if (!homeworkName) return new Response('Missing homework name query parameter.', { status: 400 });
        return handleSummaryDownload(env, homeworkName);
      }
      return new Response('Invalid method', { status: 405 });
    }

    // --- [!! NEW !!] API 路由 (用于分片上传) ---

    // API 路由 1: 创建 OneDrive 上传会话
    if (path === '/api/create-upload-session') {
      if (request.method === 'POST') {
        return handleCreateUploadSession(request, env, sessionCookie);
      }
      return new Response('Invalid method', { status: 405 });
    }

    // API 路由 2: 代理上传分片
    if (path === '/api/upload-chunk') {
      if (request.method === 'POST') {
        return handleUploadChunk(request, env, ctx, sessionCookie);
      }
      return new Response('Invalid method', { status: 405 });
    }
    
    return new Response('Not Found', { status: 404 });
  },
};

// -----------------------------------------------------------------
// 2. 认证和会话管理 (AUTH) - (无变化)
// -----------------------------------------------------------------

async function handleLogin(request, env) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const { email } = await request.json();
    
    if (!email || !email.includes('@')) {
      return new Response(JSON.stringify({ success: false, message: 'Valid email is required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const lowerEmail = email.toLowerCase();
    const emailDomain = lowerEmail.substring(lowerEmail.lastIndexOf('@')); // e.g., @school.com

    const allowedEntries = (env.TEACHER_WHITELIST || "")
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(e => e.length > 0); 

    const isAllowed = allowedEntries.includes(lowerEmail) || allowedEntries.includes(emailDomain);

    if (!isAllowed) {
      console.log(`[Auth Fail] Email: ${email}. Domain: ${emailDomain}. Not in whitelist.`);
      return new Response(JSON.stringify({ success: false, message: 'Email address or domain is not authorized.' }), {
        status: 403,
        headers: jsonHeaders,
      });
    }

    console.log(`[Auth Success] Email: ${email}.`);
    
    const code = Math.floor(100000 + Math.random() * 900000).toString();
    const kvKey = `code:${email}`;
    await env.AUTH_KV.put(kvKey, code, { expirationTtl: 300 });

    const mailSent = await sendVerificationEmail(env, email, code, env.MS_USER_ID);

    if (mailSent) {
      return new Response(JSON.stringify({ success: true, message: 'Verification code sent.' }), {
        status: 200,
        headers: jsonHeaders,
      });
    } else {
      return new Response(JSON.stringify({ success: false, message: 'Failed to send email.' }), {
        status: 500,
        headers: jsonHeaders,
      });
    }
  } catch (err) {
    console.error(`[Login Error] ${err}`);
    return new Response(JSON.stringify({ success: false, message: 'Internal server error.' }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

async function handleVerify(request, env) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const { email, code } = await request.json();
    if (!email || !code) {
      return new Response(JSON.stringify({ success: false, message: 'Email and code are required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const kvKey = `code:${email}`;
    const storedCode = await env.AUTH_KV.get(kvKey);

    if (storedCode && storedCode === code) {
      await env.AUTH_KV.delete(kvKey);

      const sessionId = crypto.randomUUID();
      const sessionKey = `session:${sessionId}`;
      // [!!] 存储 email, 而不是 session ID
      await env.AUTH_KV.put(sessionKey, email, { expirationTtl: 7 * 24 * 60 * 60 });

      const cookie = `auth-session=${sessionId}; HttpOnly; Secure; Path=/; Max-Age=${7 * 24 * 60 * 60}`;

      return new Response(JSON.stringify({ success: true, message: 'Login successful.' }), {
        status: 200,
        headers: {
          'Content-Type': 'application/json',
          'Set-Cookie': cookie,
        },
      });
    } else {
      return new Response(JSON.stringify({ success: false, message: 'Invalid or expired code.' }), {
        status: 401,
        headers: jsonHeaders,
      });
    }
  } catch (err) {
    console.error(`[Verify Error] ${err}`);
    return new Response(JSON.stringify({ success: false, message: 'Internal server error.' }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

async function handleLogout(request) {
  const cookie = 'auth-session=; HttpOnly; Secure; Path=/; Max-Age=0';
  return new Response(null, {
    status: 302, 
    headers: {
      'Set-Cookie': cookie,
      'Location': '/login', 
    },
  });
}

async function isSessionValid(kv, sessionId) {
  const sessionKey = `session:${sessionId}`;
  const email = await kv.get(sessionKey);
  return email != null;
}

function getCookie(request, name) {
  const cookieHeader = request.headers.get('Cookie');
  if (cookieHeader) {
    const cookies = cookieHeader.split(';');
    for (const cookie of cookies) {
      const [cookieName, cookieValue] = cookie.trim().split('=');
      if (cookieName === name) {
        return cookieValue;
      }
    }
  }
  return null;
}

// -----------------------------------------------------------------
// 3. 邮件发送 - (无变化)
// -----------------------------------------------------------------

async function sendVerificationEmail(env, toEmail, code, fromUserId) {
  try {
    const accessToken = await getMsGraphToken(env);
    if (!accessToken) {
      console.error('[Mail Error] Failed to get MS Graph token for sending email.');
      return false;
    }

    const appTitle = env.APP_TITLE || 'AI 批改平台';
    const emailBody = {
      message: {
        subject: `[${appTitle}] 您的登录验证码`,
        body: {
          contentType: 'HTML',
          content: `... (邮件内容同 v3) ...`,
        },
        toRecipients: [ { emailAddress: { address: toEmail } } ],
      },
      saveToSentItems: 'true',
    };

    const url = `https://graph.microsoft.com/v1.0/users/${fromUserId}/sendMail`;

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(emailBody),
    });

    if (response.status === 202) {
      console.log(`[Mail Success] Verification code sent to ${toEmail}.`);
      return true;
    } else {
      const error = await response.json();
      console.error(`[Mail Fail] Failed to send email: ${response.status}`, JSON.stringify(error, null, 2));
      return false;
    }
  } catch (err) {
    console.error(`[Mail Error] Exception: ${err}`);
    return false;
  }
}


// -----------------------------------------------------------------
// 4. [!! NEW !!] 分片上传 API
// -----------------------------------------------------------------

/**
 * POST /api/create-upload-session
 * 1. 验证文件名
 * 2. 从 MS Graph 获取 OneDrive 上传 URL
 * 3. 将 (uploadId -> oneDriveUrl) 存储在 KV 中
 */
async function handleCreateUploadSession(request, env, sessionCookie) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const { fileName } = await request.json();
    if (!fileName || !fileName.toLowerCase().endsWith('.zip')) {
      return new Response(JSON.stringify({ success: false, message: 'Invalid file name. Must be a .zip file.' }), { status: 400, headers: jsonHeaders });
    }

    const homeworkName = fileName.replace(/\.zip$/i, '');
    const basePath = env.MS_ONEDRIVE_BASE_PATH || "Apps/HomeworkGrader";
    const homeworkPath = `${basePath}/${homeworkName}`;
    const uploadPath = `${homeworkPath}/${fileName}`; // 完整路径

    const msToken = await getMsGraphToken(env);
    if (!msToken) {
      return new Response(JSON.stringify({ success: false, message: 'Failed to get auth token.' }), { status: 500, headers: jsonHeaders });
    }

    // 1. 创建 OneDrive 上传会话
    const session = await createOneDriveUploadSession(env, msToken, uploadPath);
    if (!session || !session.uploadUrl) {
      return new Response(JSON.stringify({ success: false, message: 'Failed to create OneDrive upload session.' }), { status: 500, headers: jsonHeaders });
    }

    // 2. 创建一个唯一的上传 ID, 并将其与 OneDrive URL 关联
    const uploadId = crypto.randomUUID();
    const kvKey = `upload:${uploadId}`;
    
    // 存储会话信息, 2小时过期
    const sessionData = {
      oneDriveUrl: session.uploadUrl,
      homeworkName: homeworkName,
      homeworkZipName: fileName,
      homeworkPath: homeworkPath, // e.g., Apps/HomeworkGrader/HW1
      fullZipPath: uploadPath, // e.g., Apps/HomeworkGrader/HW1/HW1.zip
    };
    await env.AUTH_KV.put(kvKey, JSON.stringify(sessionData), { expirationTtl: 7200 }); 

    console.log(`[Upload Session] Created uploadId ${uploadId} for ${fileName}`);

    // 3. 将 uploadId 返回给客户端
    return new Response(JSON.stringify({
      success: true,
      uploadId: uploadId,
      message: 'Upload session created. Ready for chunks.'
    }), { status: 200, headers: jsonHeaders });

  } catch (err) {
    console.error(`[Create Session Error] ${err}`);
    return new Response(JSON.stringify({ success: false, message: `Internal error: ${err.message}` }), { status: 500, headers: jsonHeaders });
  }
}

/**
 * POST /api/upload-chunk
 * 1. 验证 uploadId, offset, fileSize, chunkSize
 * 2. 从 KV 检索 OneDrive URL
 * 3. [!!] 将请求体 (分片) 流式传输到 OneDrive
 * 4. 检查是否为最后一个分片, 如果是, 触发后台处理
 */
async function handleUploadChunk(request, env, ctx, sessionCookie) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const url = new URL(request.url);
    const uploadId = url.searchParams.get('uploadId');
    const offset = parseInt(url.searchParams.get('offset'), 10);
    const fileSize = parseInt(url.searchParams.get('fileSize'), 10);
    const chunkSize = parseInt(request.headers.get('Content-Length'), 10); // 实际的分片大小

    if (isNaN(offset) || isNaN(fileSize) || isNaN(chunkSize) || !uploadId) {
      return new Response(JSON.stringify({ success: false, message: 'Missing or invalid query parameters (uploadId, offset, fileSize, chunkSize).' }), { status: 400, headers: jsonHeaders });
    }

    // 1. 从 KV 检索会话
    const kvKey = `upload:${uploadId}`;
    const sessionDataString = await env.AUTH_KV.get(kvKey);
    if (!sessionDataString) {
      return new Response(JSON.stringify({ success: false, message: 'Invalid or expired upload session.' }), { status: 404, headers: jsonHeaders });
    }
    const sessionData = JSON.parse(sessionDataString);
    const oneDriveUrl = sessionData.oneDriveUrl;

    // 2. 准备代理请求
    const endByte = offset + chunkSize - 1;
    const contentRange = `bytes ${offset}-${endByte}/${fileSize}`;

    console.log(`[Chunk Upload] Proxying chunk for ${uploadId}. Range: ${contentRange}`);

    // 3. [!!] 将 worker 的请求体流式传输到 OneDrive
    const oneDriveResponse = await fetch(oneDriveUrl, {
      method: 'PUT',
      headers: {
        'Content-Range': contentRange,
        'Content-Length': chunkSize.toString(),
      },
      body: request.body, // 直接传递流
      duplex: 'half', // [!!] 关键: 允许在 Worker 中流式传输请求体
    });

    if (!oneDriveResponse.ok) {
      // 检查 OneDrive 返回的错误
      const errorData = await oneDriveResponse.json();
      console.error(`[Chunk Upload Error] OneDrive failed: ${oneDriveResponse.status}`, errorData);
      return new Response(JSON.stringify({ success: false, message: `OneDrive upload error: ${errorData.error.message}` }), { status: 500, headers: jsonHeaders });
    }

    // 4. 检查是否为最后一个分片 (OneDrive 返回 201 或 200)
    if (oneDriveResponse.status === 201 || oneDriveResponse.status === 200) {
      // 上传完成!
      console.log(`[Chunk Upload] Upload complete for ${uploadId}.`);
      
      const oneDriveFile = await oneDriveResponse.json();
      const oneDriveFileId = oneDriveFile.id; // [!!] 得到文件 ID

      // 触发后台处理 (使用文件 ID)
      ctx.waitUntil(
        processHomeworkInBackground(
          env,
          oneDriveFileId, // [!!] 传递 ID, 而不是路径
          sessionData.homeworkName,
          sessionData.homeworkZipName,
          sessionData.homeworkPath
        )
      );

      // (可选) 清理 KV
      await env.AUTH_KV.delete(kvKey);

      return new Response(JSON.stringify({ success: true, status: 'complete', message: 'File upload complete. Processing started in background.' }), { status: 200, headers: jsonHeaders });
    
    } else if (oneDriveResponse.status === 202) {
      // 202 Accepted - 接收了分片, 等待更多
      const nextRange = oneDriveResponse.headers.get('nextExpectedRanges');
      const nextOffset = nextRange ? parseInt(nextRange.split('-')[0], 10) : offset + chunkSize;

      return new Response(JSON.stringify({ success: true, status: 'pending', nextExpectedOffset: nextOffset }), { status: 202, headers: jsonHeaders });
    } else {
      return new Response(JSON.stringify({ success: false, message: 'Unknown OneDrive response.' }), { status: 500, headers: jsonHeaders });
    }

  } catch (err) {
    console.error(`[Chunk Upload Error] ${err}`);
    return new Response(JSON.stringify({ success: false, message: `Internal proxy error: ${err.message}` }), { status: 500, headers: jsonHeaders });
  }
}


// -----------------------------------------------------------------
// 5. [!! REWRITTEN !!] 核心业务逻辑 (远程解压)
// -----------------------------------------------------------------

/**
 * (后台任务) 异步处理作业 (远程解压版)
 * 1. 接收 OneDrive 文件 ID
 * 2. [!!] 命令 OneDrive 远程解压主 .zip 包
 * 3. 遍历解压后的学生 .zip 包
 * 4. [!!] 命令 OneDrive 远程解压学生 .zip 包
 * 5. 遍历学生文件, 逐个下载, 构建 Gemini 请求
 * 6. 调用 Gemini, 上传批改报告
 */
async function processHomeworkInBackground(env, mainZipFileId, homeworkName, homeworkZipName, homeworkPath) {
  
  console.log(`[后台处理] 开始: ${homeworkName} (File ID: ${mainZipFileId})`);
  let msToken;

  try {
    msToken = await getMsGraphToken(env);
    if (!msToken) throw new Error('无法获取 MS Graph Token');

    // 1. (同之前) 初始化 Gemini
    const genAI = new GoogleGenerativeAI(env.GOOGLE_API_KEY);
    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
    const safetySettings = [
      { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
    ];
    
    // 2. [!! NEW !!] 命令 OneDrive 远程解压主 .zip 包
    // (这会解压到 .zip 文件所在的同一目录, 即 homeworkPath)
    console.log(`[后台处理] 正在命令 OneDrive 远程解压 (1): ${homeworkZipName}`);
    await extractOneDriveZip(env, msToken, mainZipFileId);
    
    // 3. [!! NEW !!] 列出解压后的文件夹内容
    const studentZips = await listOneDriveChildrenByPath(env, msToken, homeworkPath);

    // 4. (同之前) 遍历每个学生 .zip 包
    for (const studentItem of studentZips) {
      const studentZipName = studentItem.name;
      const studentZipId = studentItem.id;
      
      // 跳过主 .zip 包 和 非 .zip 文件
      if (!studentZipName.toLowerCase().endsWith('.zip') || studentZipName === homeworkZipName) {
        console.log(`[后台处理] 跳过: ${studentZipName} (不是学生 .zip)`);
        continue;
      }

      const studentName = studentZipName.replace(/\.zip$/i, '');
      const studentReportPath = `${homeworkPath}/${studentName}/report.json`;
      console.log(`[后台处理] 正在处理学生: ${studentName} (ID: ${studentZipId})`);
      
      try {
        // 5. [!! NEW !!] 命令 OneDrive 远程解压学生 .zip 包
        // (这会解压到 .zip 所在的同一目录, 即 homeworkPath)
        console.log(`[后台处理] 正在命令 OneDrive 远程解压 (2): ${studentZipName}`);
        await extractOneDriveZip(env, msToken, studentZipId);
        
        // 6. [!! NEW !!] 列出该学生解压后的文件
        // (注意: 解压后的文件夹名可能与 .zip 名相同, Graph API 会自动处理)
        const studentFiles = await listOneDriveChildrenByPath(env, msToken, `${homeworkPath}/${studentName}`);

        // 7. 准备 Gemini API 请求
        const promptParts = [
          '# 角色: 你是一位严格的计算机网络课程助教。',
          '# 任务: 批改这份作业。作业zip包名为 ' + homeworkName,
          '# 学生: ' + studentName,
          '# 指令: ',
          '1. 分析以下所有文件内容。',
          '2. 针对GET和HEAD方法的异同、MIME类型、HTTP报文格式等方面给出综合评价。',
          '3. 给出分数 (0-100) 和详细评语 (必须包含中文)。',
          '4. 严格按照 JSON 格式输出: {"student_name": "...", "score": 85, "feedback": "..."}',
          '---',
        ];

        let fileCount = 0;
        for (const fileItem of studentFiles) {
          if (fileItem.folder) continue; // 跳过子目录
          
          const fileName = fileItem.name;
          const fileId = fileItem.id;

          // 8. [!! NEW !!] 逐个下载小文件
          const fileContentBuffer = await downloadOneDriveFile(env, msToken, fileId);
          if (!fileContentBuffer) {
            console.log(`[后台处理] 无法下载文件: ${fileName}, 跳过`);
            continue;
          }
          
          const fileContentBase64 = bufferToBase64(fileContentBuffer);
          const mimeType = getMimeType(fileName) || 'application/octet-stream';
          
          promptParts.push(`## 文件: ${fileName} (MIME: ${mimeType}) ##`);
          promptParts.push({
            inlineData: {
              data: fileContentBase64,
              mimeType: mimeType,
            },
          });
          fileCount++;
        }
        
        if (fileCount === 0) {
          console.log(`[后台处理] 学生 ${studentName} 的 .zip 包是空的, 跳过`);
          continue;
        }

        // 9. 调用 Gemini API
        console.log(`[后台处理] 正在调用 Gemini API 批改 ${studentName}...`);
        const result = await model.generateContent({
          contents: [{ role: 'user', parts: promptParts }],
          safetySettings,
          generationConfig: {
            responseMimeType: 'application/json',
          },
        });

        const reportJsonString = result.response.text();
        
        // 10. 上传批改报告 (使用简单上传)
        await uploadFileToOneDrive_Simple(
          env,
          msToken,
          studentReportPath,
          reportJsonString,
          'application/json'
        );
        console.log(`[后台处理] ${studentName} 的批改报告上传成功: ${studentReportPath}`);

      } catch (err) {
        console.error(`[后台处理] 处理学生 ${studentName} 时出错: ${err}`);
        const errorReport = JSON.stringify({
          student_name: studentName,
          score: 0,
          feedback: `[AI 自动批改失败] 错误: ${err.message}`,
        });
        await uploadFileToOneDrive_Simple(env, msToken, studentReportPath, errorReport, 'application/json');
      }
    }
    console.log(`[后台处理] 完成: ${homeworkName}`);

    // (可选) 清理: 删除主 .zip 包和学生 .zip 包
    // ...

  } catch (err) {
    console.error(`[后台处理] 严重错误: ${err.message}`);
    // 可以在此处实现一个通知机制, 比如发一封失败邮件
  }
}

// -----------------------------------------------------------------
// 6. 成绩汇总 - (无变化)
// -----------------------------------------------------------------

async function handleSummaryDownload(env, homeworkName) {
  // (此函数逻辑与 v3 完全相同)
  // 它通过读取 Apps/HomeworkGrader/HW1/StudentA/report.json 来工作
  // 无论作业是如何处理的, 报告都在同一个地方。
  
  try {
    const basePath = env.MS_ONEDRIVE_BASE_PATH || 'Apps/HomeworkGrader';
    const homeworkPath = `${basePath}/${homeworkName}`;
    
    const msToken = await getMsGraphToken(env);
    if (!msToken) {
      return new Response('Failed to get auth token.', { status: 500 });
    }

    // 1. 列出作业文件夹下的所有内容 (即学生文件夹)
    const studentFolders = await listOneDriveChildrenByPath(env, msToken, homeworkPath);
    if (!studentFolders) {
       return new Response(`Error: Homework folder '${homeworkName}' not found.`, { status: 404 });
    }

    const reportPromises = studentFolders
      .filter(item => item.folder) // [!!] 确保只查找文件夹
      .map(folder => {
        const studentName = folder.name;
        const reportPath = `${homeworkPath}/${studentName}/report.json`;
        
        return downloadOneDriveFileByPath(env, msToken, reportPath)
          .then(buffer => {
            if (buffer) {
              const text = new TextDecoder().decode(buffer);
              return JSON.parse(text);
            }
            return { student_name: studentName, score: 'N/A', feedback: '批改未完成或 report.json 丢失' };
          })
          .catch(err => ({ student_name: studentName, score: 'N/A', feedback: `下载报告失败: ${err.message}` }));
    });

    const reports = await Promise.all(reportPromises);

    // 3. 生成 CSV
    const csvContent = generateCsv(reports);
    const safeHomeworkName = homeworkName.replace(/[^a-z0-9]/gi, '_');
    const fileName = `${safeHomeworkName}_Grades_${new Date().toISOString().split('T')[0]}.csv`;

    return new Response(csvContent, {
      headers: {
        'Content-Type': 'text/csv; charset=utf-8-sig', 
        'Content-Disposition': `attachment; filename="${fileName}"`,
      },
    });

  } catch (err) {
    console.error(`[Summary Error] ${err}`);
    return new Response(`Failed to generate summary: ${err.message}`, { status: 500 });
  }
}

function generateCsv(reports) {
  if (!reports || reports.length === 0) {
    return 'Student Name,Score,Feedback\n(No data)';
  }
  const headers = ['student_name', 'score', 'feedback'];
  let csv = '\uFEFF'; 
  csv += headers.join(',') + '\n';
  for (const report of reports) {
    const student = report.student_name || 'Unknown';
    const score = report.score !== undefined ? report.score : 'N/A';
    const feedback = report.feedback || 'No feedback';
    const cleanStudent = `"${student.replace(/"/g, '""')}"`;
    const cleanScore = `"${score.toString().replace(/"/g, '""')}"`;
    const cleanFeedback = `"${feedback.replace(/"/g, '""').replace(/\n/g, ' ')}"`;
    csv += [cleanStudent, cleanScore, cleanFeedback].join(',') + '\n';
  }
  return csv;
}


// -----------------------------------------------------------------
// 7. MS GRAPH API 辅助函数 (!! 已扩展 !!)
// -----------------------------------------------------------------

async function getMsGraphToken(env) {
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  try {
    const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body });
    if (!response.ok) { const error = await response.json(); console.error('[MS Token Error] Failed to get token:', JSON.stringify(error)); return null; }
    const data = await response.json();
    return data.access_token;
  } catch (err) { console.error(`[MS Token Error] Exception: ${err}`); return null; }
}

async function uploadFileToOneDrive_Simple(env, accessToken, pathOnOneDrive, fileContent, contentType) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/content`;
  try {
    const response = await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': contentType },
      body: fileContent,
    });
    if (response.ok) return true;
    const error = await response.json();
    console.error(`[OneDrive Error] Failed to upload (simple) ${pathOnOneDrive}: ${response.status}`, JSON.stringify(error));
    return false;
  } catch (err) { console.error(`[OneDrive Error] Upload (simple) exception: ${err}`); return false; }
}

async function createOneDriveUploadSession(env, accessToken, pathOnOneDrive) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/createUploadSession`;
  const body = { item: { '@microsoft.graph.conflictBehavior': 'replace' } };
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (response.ok) return await response.json();
    const error = await response.json();
    console.error(`[OneDrive Error] Failed to create upload session: ${response.status}`, JSON.stringify(error));
    return null;
  } catch (err) { console.error(`[OneDrive Error] createUploadSession exception: ${err}`); return null; }
}

/**
 * [!! NEW !!]
 * 命令 OneDrive 远程解压一个 .zip 文件
 */
async function extractOneDriveZip(env, accessToken, fileId) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/extract`;
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        // 可以在此指定密码, 但我们假设没有密码
        // "password": "..." 
      })
    });
    
    // 202 Accepted 意味着它已开始在后台解压
    if (response.status === 202) {
      console.log(`[OneDrive Extract] Extracting ${fileId} in background...`);
      // Graph API 不会等待解压完成, 它立即返回
      // 这在我们的流程中是可接受的
      return true;
    } 
    // (如果文件很小, 它可能会立即返回 200, 但 202 更常见)
    if (response.status === 200) {
      console.log(`[OneDrive Extract] Extract ${fileId} complete.`);
      return true;
    }
    
    const error = await response.json();
    console.error(`[OneDrive Extract] Failed to extract ${fileId}: ${response.status}`, JSON.stringify(error));
    return false;
    
  } catch (err) {
    console.error(`[OneDrive Extract] Exception: ${err}`);
    return false;
  }
}

/**
 * [!! NEW !!]
 * 按路径列出 OneDrive 文件夹的内容
 */
async function listOneDriveChildrenByPath(env, accessToken, folderPath) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}:/children?$select=id,name,folder,file`;
  try {
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` },
    });
    if (!response.ok) {
      console.error(`[OneDrive List] Failed to list children for ${folderPath}: ${response.status}`);
      return null;
    }
    const data = await response.json();
    return data.value; // 返回 items 数组
  } catch (err) {
    console.error(`[OneDrive List] Exception: ${err}`);
    return null;
  }
}

/**
 * [!! NEW !!]
 * 按 ID 下载 OneDrive 文件 (返回 ArrayBuffer)
 */
async function downloadOneDriveFile(env, accessToken, fileId) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`;
  try {
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` },
    });
    if (!response.ok) {
      console.error(`[OneDrive Download] Failed to download file ${fileId}: ${response.status}`);
      return null;
    }
    return await response.arrayBuffer();
  } catch (err) {
    console.error(`[OneDrive Download] Exception: ${err}`);
    return null;
  }
}

/**
 * [!! NEW !!]
 * 按 Path 下载 OneDrive 文件 (返回 ArrayBuffer) - 用于 /summary
 */
async function downloadOneDriveFileByPath(env, accessToken, filePath) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;
  try {
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` },
    });
    if (!response.ok) {
      // 404 (Not Found) 是正常情况, 意味着报告尚未生成
      if(response.status !== 404) {
        console.error(`[OneDrive Download] Failed to download file ${filePath}: ${response.status}`);
      }
      return null;
    }
    return await response.arrayBuffer();
  } catch (err) {
    console.error(`[OneDrive Download] Exception: ${err}`);
    return null;
  }
}


// -----------------------------------------------------------------
// 8. 工具函数 - (无变化)
// -----------------------------------------------------------------

function bufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function getMimeType(filename) {
  const extension = filename.split('.').pop().toLowerCase();
  const mimeTypes = {
    'txt': 'text/plain', 'html': 'text/html', 'css': 'text/css',
    'js': 'application/javascript', 'json': 'application/json', 'xml': 'application/xml',
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif', 'svg': 'image/svg+xml', 'pdf': 'application/pdf',
    'zip': 'application/zip', 'doc': 'application/msword',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xls': 'application/vnd.ms-excel',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'ppt': 'application/vnd.ms-powerpoint',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'pcap': 'application/vnd.tcpdump.pcap', 'pcapng': 'application/x-pcapng',
  };
  return mimeTypes[extension];
}


// -----------------------------------------------------------------
// 9. [!! REWRITTEN !!] 前端页面 (分片上传)
// -----------------------------------------------------------------

function getLoginPage(appTitle) {
  // (此函数逻辑与 v3 完全相同, 此处省略以保持简洁)
  return `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>登录 - ${appTitle}</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 flex items-center justify-center min-h-screen">
  <div class="bg-white p-8 rounded-lg shadow-md w-full max-w-md">
    <h1 class="text-2xl font-bold text-center mb-6">${appTitle}</h1>
    
    <!-- 步骤 1: 输入邮箱 -->
    <div id="step-1">
      <h2 class="text-lg font-semibold mb-4">教师登录</h2>
      <label for="email" class="block text-sm font-medium text-gray-700">邮箱地址</label>
      <input type="email" id="email" name="email" required
             class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm">
      <button id="send-code-btn" 
              class="w-full mt-4 py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500">
        发送验证码
      </button>
    </div>

    <!-- 步骤 2: 输入验证码 -->
    <div id="step-2" class="hidden">
      <h2 class="text-lg font-semibold mb-4">验证</h2>
      <p class="text-sm text-gray-600 mb-2">已发送验证码至 <strong id="email-display"></strong></p>
      <label for="code" class="block text-sm font-medium text-gray-700">6 位验证码</label>
      <input type="text" id="code" name="code" required maxlength="6"
             class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm">
      <button id="verify-btn" 
              class="w-full mt-4 py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500">
        登录
      </button>
      <button id="back-btn" 
              class="w-full mt-2 py-2 px-4 border border-gray-300 rounded-md shadow-sm text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none">
        返回
      </button>
    </div>
    
    <div id="message-box" class="mt-4 text-sm text-center"></div>
  </div>

  <script>
    const step1 = document.getElementById('step-1');
    const step2 = document.getElementById('step-2');
    const sendCodeBtn = document.getElementById('send-code-btn');
    const verifyBtn = document.getElementById('verify-btn');
    const backBtn = document.getElementById('back-btn');
    const emailInput = document.getElementById('email');
    const codeInput = document.getElementById('code');
    const emailDisplay = document.getElementById('email-display');
    const messageBox = document.getElementById('message-box');
    let currentEmail = '';

    sendCodeBtn.onclick = async () => {
      currentEmail = emailInput.value;
      if (!currentEmail) { showMessage('请输入邮箱地址。', 'red'); return; }
      showMessage('正在发送验证码...', 'blue');
      sendCodeBtn.disabled = true; sendCodeBtn.innerText = '发送中...';
      try {
        const response = await fetch('/login', {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: currentEmail }),
        });
        const result = await response.json();
        if (response.ok && result.success) {
          showMessage('验证码已发送, 请查收邮件。', 'green');
          emailDisplay.innerText = currentEmail;
          step1.classList.add('hidden'); step2.classList.remove('hidden');
        } else {
          showMessage(result.message || '发送失败', 'red');
        }
      } catch (err) { showMessage('请求失败: ' + err.message, 'red');
      } finally { sendCodeBtn.disabled = false; sendCodeBtn.innerText = '发送验证码'; }
    };
    
    verifyBtn.onclick = async () => {
      const code = codeInput.value;
      if (!code || code.length !== 6) { showMessage('请输入 6 位验证码。', 'red'); return; }
      showMessage('正在登录...', 'blue');
      verifyBtn.disabled = true; verifyBtn.innerText = '登录中...';
      try {
        const response = await fetch('/verify', {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: currentEmail, code: code }),
        });
        const result = await response.json();
        if (response.ok && result.success) {
          showMessage('登录成功！正在跳转...', 'green');
          window.location.href = '/'; 
        } else {
          showMessage(result.message || '验证失败', 'red');
          verifyBtn.disabled = false; verifyBtn.innerText = '登录';
        }
      } catch (err) {
        showMessage('请求失败: ' + err.message, 'red');
        verifyBtn.disabled = false; verifyBtn.innerText = '登录';
      }
    };
    backBtn.onclick = () => {
      step1.classList.remove('hidden'); step2.classList.add('hidden');
      messageBox.innerHTML = ''; currentEmail = '';
    };
    function showMessage(message, color) { messageBox.innerHTML = \`<span class="text-\${color}-600">\${message}</span>\`; }
  </script>
</body>
</html>
  `;
}


function getHtmlPage(appTitle) {
  // [!!] 这是 v4 的前端, 实现了分片上传逻辑
  return `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${appTitle}</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 min-h-screen">

  <nav class="bg-white shadow-md">
    <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex justify-between h-16">
        <div class="flex items-center"><span class="text-2xl font-bold text-gray-800">${appTitle}</span></div>
        <div class="flex items-center"><a href="/logout" class="text-sm font-medium text-gray-600 hover:text-gray-900">退出登录</a></div>
      </div>
    </div>
  </nav>

  <div class="max-w-4xl mx-auto p-4 sm:p-6 lg:p-8">
    
    <!-- 上传区 -->
    <div class="bg-white p-6 rounded-lg shadow-md mb-6">
      <h2 class="text-xl font-semibold mb-4">1. 上传作业压缩包 (无大小限制)</h2>
      <p class="text-sm text-gray-600 mb-4">
        请上传 .zip 压缩包。系统将使用分片上传, 支持大于 100MB 的文件。
      </p>
      
      <!-- [!!] v4: 这是一个 JS 驱动的表单, 不是标准 form -->
      <div id="upload-form">
        <label for="zipfile" class="block text-sm font-medium text-gray-700">选择 .zip 文件</label>
        <input type="file" id="zipfile" name="zipfile" accept=".zip" required
               class="mt-1 block w-full text-sm text-gray-500
                      file:mr-4 file:py-2 file:px-4
                      file:rounded-md file:border-0
                      file:text-sm file:font-semibold
                      file:bg-blue-50 file:text-blue-700
                      hover:file:bg-blue-100">
        
        <button type="button" id="upload-btn" 
                class="w-full mt-4 py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500">
          上传并开始后台批改
        </button>
      </div>
      
      <!-- [!!] v4: 进度条 -->
      <div id="progress-container" class="w-full bg-gray-200 rounded-full h-2.5 mt-4 hidden">
        <div id="progress-bar" class="bg-blue-600 h-2.5 rounded-full" style="width: 0%"></div>
      </div>
      <div id="upload-message" class="mt-4 text-sm"></div>
    </div>

    <!-- 下载区 -->
    <div class="bg-white p-6 rounded-lg shadow-md">
      <h2 class="text-xl font-semibold mb-4">2. 下载成绩总表 (CSV)</h2>
      <p class="text-sm text-gray-600 mb-4">
        AI 批改需要几分钟时间。请在上传几分钟后, 在此下载成绩汇总表。
      </p>
      <form id="download-form">
        <label for="homework-name" class="block text-sm font-medium text-gray-700">
          作业名称
          <span class="text-xs text-gray-500">(必须与您上传的 .zip 文件名(不含.zip)完全一致)</span>
        </label>
        <input type="text" id="homework-name" name="homework-name" required
               class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
               placeholder="例如: 22计网1-GET+HEAD(附件)">
        <button type="submit" id="download-btn"
                class="w-full mt-4 py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500">
          下载成绩总表 (.csv)
        </button>
      </form>
      <div id="download-message" class="mt-4 text-sm"></div>
    </div>
  </div>

  <script>
    // --- [!!] v4: 分片上传 (Chunking Upload) ---
    const CHUNK_SIZE = 10 * 1024 * 1024; // 10 MB - 必须小于 Worker 的 100MB 限制
    const fileInput = document.getElementById('zipfile');
    const uploadBtn = document.getElementById('upload-btn');
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const uploadMessage = document.getElementById('upload-message');

    uploadBtn.onclick = async () => {
      if (!fileInput.files || fileInput.files.length === 0) {
        showMessage(uploadMessage, '请先选择一个 .zip 文件。', 'red');
        return;
      }
      const file = fileInput.files[0];
      if (!file.name.toLowerCase().endsWith('.zip')) {
        showMessage(uploadMessage, '必须是 .zip 文件。', 'red');
        return;
      }

      uploadBtn.disabled = true;
      uploadBtn.innerText = '正在初始化...';
      progressContainer.classList.remove('hidden');
      updateProgress(0);

      try {
        // 1. 创建上传会话
        showMessage(uploadMessage, '步骤 1/3: 正在创建 OneDrive 上传会话...', 'blue');
        const sessionResponse = await fetch('/api/create-upload-session', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fileName: file.name }),
        });
        
        const sessionData = await sessionResponse.json();
        if (!sessionResponse.ok || !sessionData.success) {
          throw new Error(sessionData.message || '无法创建上传会话。');
        }
        
        const { uploadId } = sessionData;
        
        // 2. 开始分片上传
        showMessage(uploadMessage, '步骤 2/3: 正在上传文件分片...', 'blue');
        await uploadFileInChunks(file, uploadId);

        // 3. 完成
        // (注意: 'complete' 消息由最后一个 chunk 处理器在后台触发)
        showMessage(uploadMessage, \`<strong>成功:</strong> 文件 "\${file.name}" 已上传, 正在后台处理批改。这可能需要几分钟时间。\`, 'green');
        fileInput.value = ''; // 清空
        
      } catch (err) {
        showMessage(uploadMessage, \`<strong>上传失败:</strong> \${err.message}\`, 'red');
      } finally {
        uploadBtn.disabled = false;
        uploadBtn.innerText = '上传并开始后台批改';
        // 5秒后隐藏进度条
        setTimeout(() => {
           progressContainer.classList.add('hidden');
           updateProgress(0);
        }, 5000);
      }
    };

    async function uploadFileInChunks(file, uploadId) {
      const fileSize = file.size;
      let offset = 0;
      
      while (offset < fileSize) {
        const chunkEnd = Math.min(offset + CHUNK_SIZE, fileSize);
        const chunk = file.slice(offset, chunkEnd);
        const chunkSize = chunk.size;
        
        const uploadUrl = \`/api/upload-chunk?uploadId=\${uploadId}&offset=\${offset}&fileSize=\${fileSize}\`;
        
        // 更新 UI
        const percent = Math.round((offset / fileSize) * 100);
        updateProgress(percent);
        showMessage(uploadMessage, \`步骤 2/3: 正在上传分片... (\${(offset / (1024*1024)).toFixed(1)} MB / \${(fileSize / (1024*1024)).toFixed(1)} MB)\`, 'blue');

        const response = await fetch(uploadUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/octet-stream', // 发送原始字节
            'Content-Length': chunkSize.toString(), // [!!] 告知 Worker (代理) 分片大小
          },
          body: chunk,
        });

        const result = await response.json();
        
        if (!response.ok || !result.success) {
          throw new Error(result.message || \`分片上传失败, 偏移量: \${offset}\`);
        }
        
        if (result.status === 'complete') {
          // 最后一个分片已上传
          updateProgress(100);
          return;
        }

        if (result.nextExpectedOffset) {
          offset = result.nextExpectedOffset;
        } else {
          // 备用
          offset = chunkEnd;
        }
      }
      updateProgress(100);
    }
    
    function updateProgress(percent) {
      progressBar.style.width = \`\${percent}%\`;
    }

    // --- 下载逻辑 (无变化) ---
    const downloadForm = document.getElementById('download-form');
    const downloadBtn = document.getElementById('download-btn');
    const downloadMessage = document.getElementById('download-message');
    const homeworkNameInput = document.getElementById('homework-name');

    downloadForm.onsubmit = async (e) => {
      e.preventDefault();
      const homeworkName = homeworkNameInput.value.trim();
      if (!homeworkName) { showMessage(downloadMessage, '请输入作业名称。', 'red'); return; }
      downloadBtn.disabled = true; downloadBtn.innerText = '正在生成...';
      showMessage(downloadMessage, '正在请求成绩总表, 请稍候...', 'blue');

      try {
        const url = \`/summary?homework=\${encodeURIComponent(homeworkName)}\`;
        const response = await fetch(url, { method: 'GET' });

        if (response.ok) {
          const blob = await response.blob();
          const contentDisposition = response.headers.get('content-disposition');
          let filename = 'grades.csv';
          if (contentDisposition) {
            const match = contentDisposition.match(/filename="?([^"]+)"?/);
            if (match) filename = match[1];
          }
          const link = document.createElement('a');
          link.href = URL.createObjectURL(blob);
          link.download = filename;
          document.body.appendChild(link); link.click(); document.body.removeChild(link);
          showMessage(downloadMessage, '下载已开始!', 'green');
        } else {
          const text = await response.text();
          showMessage(downloadMessage, \`<strong>错误 (\${response.status}):</strong> \${text}\`, 'red');
        }
      } catch (err) {
        showMessage(downloadMessage, \`<strong>请求失败:</strong> \${err.message}\`, 'red');
      } finally {
        downloadBtn.disabled = false;
        downloadBtn.innerText = '下载成绩总表 (.csv)';
      }
    };
    
    function showMessage(element, message, color) {
      element.innerHTML = \`<span class="text-\${color}-600">\${message}</span>\`;
    }
  </script>
</body>
</html>
  `;
}

