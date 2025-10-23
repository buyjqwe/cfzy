/**
 * Cloudflare Worker - AI 自动作业批改平台
 *
 * 功能:
 * 1. 教师邮箱验证码登录 (支持邮箱或域名白名单)
 * 2. 异步后台处理作业:
 * - 接收 .zip 包
 * - 上传 .zip 到 OneDrive
 * - (后台) 解压, 逐个学生调用 Gemini API 批改
 * - (后台) 将 report.json 上传回 OneDrive
 * 3. 教师下载 CSV 成绩总表
 *
 * 依赖:
 * - KV (AUTH_KV): 存储登录验证码
 * - Secrets:
 * - GOOGLE_API_KEY
 * - MS_CLIENT_ID
 * - MS_CLIENT_SECRET
 * - MS_TENANT_ID
 * - MS_USER_ID (用于 Graph API 发件)
 * - MS_ONEDRIVE_BASE_PATH
 * - APP_TITLE
 * - TEACHER_WHITELIST (逗号分隔的邮箱或域名, e.g., "admin@school.com,@school.com")
 */

import { unzip } from 'unzipit'; // 用于解压 .zip 文件
import { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } from '@google/generative-ai'; // Gemini API

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

    // 登录页路由
    if (path === '/login') {
      if (request.method === 'POST') {
        return handleLogin(request, env);
      }
      return new Response(getLoginPage(env.APP_TITLE || "AI 批改平台"), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }

    // 验证码验证路由
    if (path === '/verify') {
      if (request.method === 'POST') {
        return handleVerify(request, env);
      }
      return new Response('Invalid method', { status: 405 });
    }
    
    // 登出路由
    if (path === '/logout') {
      return handleLogout(request);
    }

    // --- 以下路由需要认证 ---
    if (!isAuthenticated) {
      // 未认证, 重定向到登录页
      return Response.redirect(new URL('/login', request.url).toString(), 302);
    }

    // 根路径 (上传页面)
    if (path === '/') {
      if (request.method === 'POST') {
        // (异步) 处理上传
        return handleUpload(request, env, ctx);
      }
      // (GET) 显示主页
      return new Response(getHtmlPage(env.APP_TITLE || "AI 批改平台"), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }

    // 成绩汇总下载路由
    if (path === '/summary') {
      if (request.method === 'GET') {
        const homeworkName = url.searchParams.get('homework');
        if (!homeworkName) {
          return new Response('Missing homework name query parameter.', { status: 400 });
        }
        return handleSummaryDownload(env, homeworkName);
      }
      return new Response('Invalid method', { status: 405 });
    }
    
    // 默认 404
    return new Response('Not Found', { status: 404 });
  },
};

// -----------------------------------------------------------------
// 2. 认证和会话管理 (AUTH)
// -----------------------------------------------------------------

/**
 * 处理登录请求 (POST /login)
 * 1. 验证邮箱是否在白名单 (或白名单域名)
 * 2. 生成 6 位验证码
 * 3. 存储验证码到 KV (5分钟过期)
 * 4. 发送验证码邮件
 */
async function handleLogin(request, env) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const { email } = await request.json();
    
    // [!!! 已更新 !!!] 
    // 支持邮箱全名或 @domain.com 格式的白名单
    if (!email || !email.includes('@')) {
      // [BUG FIX]: 总是返回 JSON
      return new Response(JSON.stringify({ success: false, message: 'Valid email is required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const lowerEmail = email.toLowerCase();
    const emailDomain = lowerEmail.substring(lowerEmail.lastIndexOf('@')); // e.g., @school.com

    // 验证邮箱或邮箱域是否在白名单中
    const allowedEntries = (env.TEACHER_WHITELIST || "")
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(e => e.length > 0); // 过滤掉空字符串

    // 检查邮箱全名 或 邮箱域名是否在允许列表中
    const isAllowed = allowedEntries.includes(lowerEmail) || allowedEntries.includes(emailDomain);

    if (!isAllowed) {
      console.log(`[Auth Fail] Email: ${email}. Domain: ${emailDomain}. Not in whitelist.`);
      // [BUG FIX]: 总是返回 JSON
      return new Response(JSON.stringify({ success: false, message: 'Email address or domain is not authorized.' }), {
        status: 403,
        headers: jsonHeaders,
      });
    }
    // [!!! 更新结束 !!!]

    console.log(`[Auth Success] Email: ${email}.`);
    
    // 生成 6 位验证码
    const code = Math.floor(100000 + Math.random() * 900000).toString();
    const kvKey = `code:${email}`;
    // 存储验证码, 5分钟 (300秒) 过期
    await env.AUTH_KV.put(kvKey, code, { expirationTtl: 300 });

    // 发送邮件 (不阻塞响应)
    // 注意: MS_USER_ID 是指 *发件人* 的邮箱地址或用户 ID
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
    // [BUG FIX]: 总是返回 JSON
    return new Response(JSON.stringify({ success: false, message: 'Internal server error.' }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

/**
 * 处理验证码 (POST /verify)
 * 1. 检查验证码是否匹配
 * 2. 创建会话 (Session)
 * 3. 存储会话到 KV (7天过期)
 * 4. 设置 HttpOnly Cookie
 */
async function handleVerify(request, env) {
  const jsonHeaders = { 'Content-Type': 'application/json' };
  try {
    const { email, code } = await request.json();
    if (!email || !code) {
      // [BUG FIX]: 总是返回 JSON
      return new Response(JSON.stringify({ success: false, message: 'Email and code are required.' }), {
        status: 400,
        headers: jsonHeaders,
      });
    }

    const kvKey = `code:${email}`;
    const storedCode = await env.AUTH_KV.get(kvKey);

    if (storedCode && storedCode === code) {
      // 验证成功, 删除验证码
      await env.AUTH_KV.delete(kvKey);

      // 创建会话
      const sessionId = crypto.randomUUID();
      const sessionKey = `session:${sessionId}`;
      // 存储会话, 7天过期
      await env.AUTH_KV.put(sessionKey, email, { expirationTtl: 7 * 24 * 60 * 60 });

      // 设置 HttpOnly Cookie
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
    // [BUG FIX]: 总是返回 JSON
    return new Response(JSON.stringify({ success: false, message: 'Internal server error.' }), {
      status: 500,
      headers: jsonHeaders,
    });
  }
}

/**
 * 处理登出 (GET /logout)
 * 删除 Cookie
 */
async function handleLogout(request) {
  // 设置一个立即过期的 Cookie 来删除它
  const cookie = 'auth-session=; HttpOnly; Secure; Path=/; Max-Age=0';
  return new Response(null, {
    status: 302, // 重定向
    headers: {
      'Set-Cookie': cookie,
      'Location': '/login', // 重定向到登录页
    },
  });
}

/**
 * 验证会话 ID 是否有效
 */
async function isSessionValid(kv, sessionId) {
  const sessionKey = `session:${sessionId}`;
  const email = await kv.get(sessionKey);
  return email != null;
}

/**
 * 从请求中解析 Cookie
 */
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
// 3. 邮件发送 (MS GRAPH)
// -----------------------------------------------------------------

/**
 * 使用 MS Graph API 发送验证码邮件
 * (需要 Mail.Send 应用程序权限)
 */
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
          content: `
            <div style="font-family: Arial, sans-serif; line-height: 1.6;">
              <h2>${appTitle}</h2>
              <p>您好,</p>
              <p>您的登录验证码是:</p>
              <h1 style="color: #333; letter-spacing: 2px;">${code}</h1>
              <p>此验证码 5 分钟内有效。</p>
              <p>如果您没有请求此验证码, 请忽略此邮件。</p>
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

    // MS_USER_ID 是指发件人的 User ID 或 UPN (e.g., your-email@example.com)
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
// 4. 核心业务逻辑 (作业处理)
// -----------------------------------------------------------------

/**
 * 处理文件上传 (POST /)
 * 1. 立即响应 "Accepted"
 * 2. 使用 ctx.waitUntil 在后台执行耗时任务
 */
async function handleUpload(request, env, ctx) {
  try {
    const formData = await request.formData();
    const file = formData.get('zipfile');

    // [BUG FIX]: 检查是否真的上传了文件
    if (!file || typeof file === 'string' || !file.name) {
      return new Response('No file uploaded or invalid file format. Please select a .zip file.', { status: 400 });
    }

    const homeworkZipName = file.name;
    const fileBuffer = await file.arrayBuffer();

    // 1. 立即响应客户端, 防止浏览器超时
    const responseMessage = `文件 "${homeworkZipName}" 已收到，正在后台处理批改。这可能需要几分钟时间。`;
    
    // 2. 告诉 Worker 在响应关闭后继续执行
    ctx.waitUntil(
      processHomeworkInBackground(env, fileBuffer, homeworkZipName)
    );

    // 3. 返回 202 Accepted
    return new Response(responseMessage, { status: 202 });

  } catch (err) {
    console.error(`[Upload Error] ${err}`);
    return new Response(`Error processing upload: ${err.message}`, { status: 500 });
  }
}

/**
 * (后台任务) 异步处理作业
 * 1. 上传主 .zip 包到 OneDrive
 * 2. 解压主包
 * 3. 遍历学生包, 调用 Gemini
 * 4. 上传批改报告
 */
async function processHomeworkInBackground(env, fileBuffer, homeworkZipName) {
  const homeworkName = homeworkZipName.replace(/\.zip$/i, ''); // e.g., "22计网1-GET+HEAD(附件)"
  const basePath = env.MS_ONEDRIVE_BASE_PATH || "Apps/HomeworkGrader";
  const homeworkPath = `${basePath}/${homeworkName}`; // e.g., Apps/HomeworkGrader/22计网1-GET+HEAD(附件)
  
  console.log(`[后台处理] 开始: ${homeworkZipName}`);

  try {
    const msToken = await getMsGraphToken(env);
    if (!msToken) {
      console.error('[后台处理] 失败: 无法获取 MS Graph Token');
      return;
    }

    // 1. 上传原始 .zip 包
    const uploadPath = `${homeworkPath}/${homeworkZipName}`;
    const uploadSuccess = await uploadFileToOneDrive(env, msToken, uploadPath, fileBuffer, 'application/zip');
    if (!uploadSuccess) {
      console.error(`[后台处理] 失败: 上传主 .zip 包到 ${uploadPath} 失败`);
      return;
    }
    console.log(`[后台处理] 主 .zip 包上传成功: ${uploadPath}`);

    // 2. 初始化 Gemini 模型
    const genAI = new GoogleGenerativeAI(env.GOOGLE_API_KEY);
    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
    const safetySettings = [
      { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
      { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
    ];

    // 3. 解压主 .zip 包 (在内存中)
    const { entries: studentZips } = await unzip(fileBuffer);
    
    // 4. 遍历每个学生 .zip 包
    for (const [studentZipName, studentZipFile] of Object.entries(studentZips)) {
      if (studentZipFile.isDirectory || !studentZipName.toLowerCase().endsWith('.zip')) {
        console.log(`[后台处理] 跳过: ${studentZipName} (不是 .zip 文件)`);
        continue;
      }

      const studentName = studentZipName.replace(/\.zip$/i, '').split('/').pop(); // e.g., "学生A"
      console.log(`[后台处理] 正在处理学生: ${studentName}`);
      
      try {
        const studentZipBuffer = await studentZipFile.arrayBuffer();
        
        // 5. 解压学生 .zip 包 (在内存中)
        const { entries: studentFiles } = await unzip(studentZipBuffer);

        // 6. 准备 Gemini API 请求
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
        for (const [fileName, fileData] of Object.entries(studentFiles)) {
          if (fileData.isDirectory) continue;
          
          const fileContentBuffer = await fileData.arrayBuffer();
          const fileContentBase64 = bufferToBase64(fileContentBuffer);
          const mimeType = getMimeType(fileName) || 'application/octet-stream';
          
          // 添加文件内容到 prompt
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

        // 7. 调用 Gemini API
        console.log(`[后台处理] 正在调用 Gemini API 批改 ${studentName}...`);
        const result = await model.generateContent({
          contents: [{ role: 'user', parts: promptParts }],
          safetySettings,
          generationConfig: {
            responseMimeType: 'application/json',
          },
        });

        const reportJsonString = result.response.text();
        
        // 8. 上传批改报告
        const reportPath = `${homeworkPath}/${studentName}/report.json`;
        await uploadFileToOneDrive(
          env,
          msToken,
          reportPath,
          reportJsonString,
          'application/json'
        );
        console.log(`[后台处理] ${studentName} 的批改报告上传成功: ${reportPath}`);

      } catch (err) {
        console.error(`[后台处理] 处理学生 ${studentName} 时出错: ${err}`);
        // 尝试上传一个错误报告
        const errorReport = JSON.stringify({
          student_name: studentName,
          score: 0,
          feedback: `[AI 自动批改失败] 错误: ${err.message}`,
        });
        const reportPath = `${homeworkPath}/${studentName}/report.json`;
        await uploadFileToOneDrive(env, msToken, reportPath, errorReport, 'application/json');
      }
    }
    console.log(`[后台处理] 完成: ${homeworkZipName}`);
  } catch (err) {
    console.error(`[后台处理] 严重错误: ${err}`);
    // 可以在此处添加逻辑, 比如上传一个总的 ERROR.txt 文件到 OneDrive
  }
}

/**
 * 处理成绩总表下载 (GET /summary)
 * 1. 遍历 OneDrive 上的学生文件夹
 * 2. 下载每个 report.json
 * 3. 合并成 CSV
 */
async function handleSummaryDownload(env, homeworkName) {
  try {
    const basePath = env.MS_ONEDRIVE_BASE_PATH || 'Apps/HomeworkGrader';
    const homeworkPath = `${basePath}/${homeworkName}`;
    
    const msToken = await getMsGraphToken(env);
    if (!msToken) {
      return new Response('Failed to get auth token.', { status: 500 });
    }

    // 1. 列出作业文件夹下的所有内容 (即学生文件夹)
    const listUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${homeworkPath}:/children?$select=name,folder`;
    const listResponse = await fetch(listUrl, {
      headers: { 'Authorization': `Bearer ${msToken}` },
    });

    if (!listResponse.ok) {
      if (listResponse.status === 404) {
        return new Response(`Error: Homework folder '${homeworkName}' not found.`, { status: 404 });
      }
      return new Response('Failed to list homework contents.', { status: listResponse.status });
    }

    const { value: items } = await listResponse.json();
    const studentFolders = items.filter(item => item.folder); // 只保留文件夹

    if (studentFolders.length === 0) {
      return new Response(`No student submissions found in '${homeworkName}'.`, { status: 404 });
    }

    // 2. 并发下载所有 report.json
    const reportPromises = studentFolders.map(folder => {
      const studentName = folder.name;
      const reportPath = `${homeworkPath}/${studentName}/report.json`;
      const reportUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${reportPath}:/content`;
      
      return fetch(reportUrl, { headers: { 'Authorization': `Bearer ${msToken}` } })
        .then(res => {
          if (res.ok) return res.json();
          // 如果报告不存在 (404), 返回一个表示未完成的对象
          return { student_name: studentName, score: 'N/A', feedback: '批改未完成或 report.json 丢失' };
        })
        .catch(err => ({ student_name: studentName, score: 'N/A', feedback: `下载报告失败: ${err.message}` }));
    });

    const reports = await Promise.all(reportPromises);

    // 3. 生成 CSV
    const csvContent = generateCsv(reports);
    const safeHomeworkName = homeworkName.replace(/[^a-z0-9]/gi, '_'); // 清理文件名
    const fileName = `${safeHomeworkName}_Grades_${new Date().toISOString().split('T')[0]}.csv`;

    return new Response(csvContent, {
      headers: {
        'Content-Type': 'text/csv; charset=utf-8-sig', // utf-8-sig 确保 Excel 正确打开中文
        'Content-Disposition': `attachment; filename="${fileName}"`,
      },
    });

  } catch (err) {
    console.error(`[Summary Error] ${err}`);
    return new Response(`Failed to generate summary: ${err.message}`, { status: 500 });
  }
}

/**
 * 将 JSON 报告数组转换为 CSV 字符串
 */
function generateCsv(reports) {
  if (!reports || reports.length === 0) {
    return 'Student Name,Score,Feedback\n(No data)';
  }

  const headers = ['student_name', 'score', 'feedback'];
  // 添加 UTF-8 BOM, 帮助 Excel 正确识别编码
  let csv = '\uFEFF'; 
  csv += headers.join(',') + '\n';

  for (const report of reports) {
    const student = report.student_name || 'Unknown';
    const score = report.score !== undefined ? report.score : 'N/A';
    const feedback = report.feedback || 'No feedback';
    
    // 清理数据, 避免 CSV 注入和换行符问题
    const cleanStudent = `"${student.replace(/"/g, '""')}"`;
    const cleanScore = `"${score.toString().replace(/"/g, '""')}"`;
    const cleanFeedback = `"${feedback.replace(/"/g, '""').replace(/\n/g, ' ')}"`; // 移除换行符
    
    csv += [cleanStudent, cleanScore, cleanFeedback].join(',') + '\n';
  }
  return csv;
}

// -----------------------------------------------------------------
// 5. MS GRAPH API 辅助函数
// -----------------------------------------------------------------

/**
 * 获取 MS Graph API Token (客户端凭据流)
 */
async function getMsGraphToken(env) {
  // 注意: 在 Worker 中, 我们可以简单地缓存 token
  // 但对于高可用性, 更好的做法是使用 KV 存储带 TTL 的 token
  // 这里为了简单起见, 暂时不缓存
  
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: body,
    });

    if (!response.ok) {
      const error = await response.json();
      console.error('[MS Token Error] Failed to get token:', JSON.stringify(error));
      return null;
    }

    const data = await response.json();
    return data.access_token;
  } catch (err) {
    console.error(`[MS Token Error] Exception: ${err}`);
    return null;
  }
}

/**
 * 上传文件到 OneDrive
 * (使用 PUT session upload 适合大文件, 但这里为了简单, 使用 PUT < 4MB)
 */
async function uploadFileToOneDrive(env, accessToken, pathOnOneDrive, fileContent, contentType) {
  // PUT 方法适用于 < 4MB 的文件
  // 对于 > 4MB, 需要使用 uploadSession
  // 假设 report.json 和主 .zip (在 CF Worker 限制下) < 4MB
  
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${pathOnOneDrive}:/content`;
  
  try {
    const response = await fetch(url, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': contentType,
      },
      body: fileContent, // fileContent 可以是 ArrayBuffer 或 String
    });

    if (response.ok) {
      // console.log(`[OneDrive] File uploaded: ${pathOnOneDrive}`);
      return true;
    } else {
      const error = await response.json();
      console.error(`[OneDrive Error] Failed to upload ${pathOnOneDrive}: ${response.status}`, JSON.stringify(error));
      return false;
    }
  } catch (err) {
    console.error(`[OneDrive Error] Upload exception: ${err}`);
    return false;
  }
}

// -----------------------------------------------------------------
// 6. 工具函数
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
    'txt': 'text/plain',
    'html': 'text/html',
    'css': 'text/css',
    'js': 'application/javascript',
    'json': 'application/json',
    'xml': 'application/xml',
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'svg': 'image/svg+xml',
    'pdf': 'application/pdf',
    'zip': 'application/zip',
    'doc': 'application/msword',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xls': 'application/vnd.ms-excel',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'ppt': 'application/vnd.ms-powerpoint',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'pcap': 'application/vnd.tcpdump.pcap',
    'pcapng': 'application/x-pcapng',
    // 更多...
  };
  return mimeTypes[extension];
}

// -----------------------------------------------------------------
// 7. 前端页面 (HTML/CSS/JS)
// -----------------------------------------------------------------

/**
 * 登录页面
 */
function getLoginPage(appTitle) {
  // 使用 Tailwind CSS CDN
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
      if (!currentEmail) {
        showMessage('请输入邮箱地址。', 'red');
        return;
      }
      
      showMessage('正在发送验证码...', 'blue');
      sendCodeBtn.disabled = true;
      sendCodeBtn.innerText = '发送中...';

      try {
        const response = await fetch('/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email: currentEmail }),
        });
        
        const result = await response.json();
        
        if (response.ok && result.success) {
          showMessage('验证码已发送, 请查收邮件。', 'green');
          emailDisplay.innerText = currentEmail;
          step1.classList.add('hidden');
          step2.classList.remove('hidden');
        } else {
          showMessage(result.message || '发送失败: ' + response.statusText, 'red');
        }
      } catch (err) {
        showMessage('请求失败: ' + err.message, 'red');
      } finally {
        sendCodeBtn.disabled = false;
        sendCodeBtn.innerText = '发送验证码';
      }
    };
    
    verifyBtn.onclick = async () => {
      const code = codeInput.value;
      if (!code || code.length !== 6) {
        showMessage('请输入 6 位验证码。', 'red');
        return;
      }

      showMessage('正在登录...', 'blue');
      verifyBtn.disabled = true;
      verifyBtn.innerText = '登录中...';

      try {
        const response = await fetch('/verify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email: currentEmail, code: code }),
        });
        
        const result = await response.json();
        
        if (response.ok && result.success) {
          showMessage('登录成功！正在跳转...', 'green');
          // 登录成功, 服务器会设置 Cookie, 刷新页面即可
          window.location.href = '/'; 
        } else {
          showMessage(result.message || '验证失败: ' + response.statusText, 'red');
          verifyBtn.disabled = false;
          verifyBtn.innerText = '登录';
        }
      } catch (err) {
        showMessage('请求失败: ' + err.message, 'red');
        verifyBtn.disabled = false;
        verifyBtn.innerText = '登录';
      }
    };

    backBtn.onclick = () => {
      step1.classList.remove('hidden');
      step2.classList.add('hidden');
      messageBox.innerHTML = '';
      currentEmail = '';
    };

    function showMessage(message, color) {
      messageBox.innerHTML = \`<span class="text-\${color}-600">\${message}</span>\`;
    }
  </script>
</body>
</html>
  `;
}

/**
 * 主应用页面 (上传 / 下载)
 */
function getHtmlPage(appTitle) {
  // 使用 Tailwind CSS CDN
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

  <!-- 导航栏 -->
  <nav class="bg-white shadow-md">
    <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex justify-between h-16">
        <div class="flex items-center">
          <span class="text-2xl font-bold text-gray-800">${appTitle}</span>
        </div>
        <div class="flex items-center">
          <a href="/logout" class="text-sm font-medium text-gray-600 hover:text-gray-900">退出登录</a>
        </div>
      </div>
    </div>
  </nav>

  <!-- 主内容区 -->
  <div class="max-w-4xl mx-auto p-4 sm:p-6 lg:p-8">
    
    <!-- 上传区 -->
    <div class="bg-white p-6 rounded-lg shadow-md mb-6">
      <h2 class="text-xl font-semibold mb-4">1. 上传作业压缩包</h2>
      <p class="text-sm text-gray-600 mb-4">
        请上传一个 .zip 压缩包。该压缩包内应包含所有学生的作业 (每个学生一个 .zip)。<br>
        例如: "22计网1作业.zip" 内部包含 "学生A.zip", "学生B.zip" ...
      </p>
      
      <form id="upload-form" action="/" method="POST" enctype="multipart/form-data">
        <label for="zipfile" class="block text-sm font-medium text-gray-700">选择 .zip 文件</label>
        <input type="file" id="zipfile" name="zipfile" accept=".zip" required
               class="mt-1 block w-full text-sm text-gray-500
                      file:mr-4 file:py-2 file:px-4
                      file:rounded-md file:border-0
                      file:text-sm file:font-semibold
                      file:bg-blue-50 file:text-blue-700
                      hover:file:bg-blue-100">
        
        <button type="submit" id="upload-btn" 
                class="w-full mt-4 py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500">
          上传并开始后台批改
        </button>
      </form>
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
    // --- 上传逻辑 ---
    const uploadForm = document.getElementById('upload-form');
    const uploadBtn = document.getElementById('upload-btn');
    const uploadMessage = document.getElementById('upload-message');

    uploadForm.onsubmit = async (e) => {
      e.preventDefault();
      
      const formData = new FormData(uploadForm);
      const fileInput = document.getElementById('zipfile');
      
      if (!fileInput.files || fileInput.files.length === 0) {
        showMessage(uploadMessage, '请先选择一个 .zip 文件。', 'red');
        return;
      }

      uploadBtn.disabled = true;
      uploadBtn.innerText = '上传中...';
      showMessage(uploadMessage, '正在上传文件, 请勿关闭页面...', 'blue');

      try {
        const response = await fetch('/', {
          method: 'POST',
          body: formData,
        });

        const text = await response.text();

        if (response.status === 202) { // Accepted
          showMessage(uploadMessage, \`<strong>成功:</strong> \${text}\`, 'green');
        } else {
          showMessage(uploadMessage, \`<strong>错误 (\${response.status}):</strong> \${text}\`, 'red');
        }
      } catch (err) {
        showMessage(uploadMessage, \`<strong>请求失败:</strong> \${err.message}\`, 'red');
      } finally {
        uploadBtn.disabled = false;
        uploadBtn.innerText = '上传并开始后台批改';
        uploadForm.reset();
      }
    };

    // --- 下载逻辑 ---
    const downloadForm = document.getElementById('download-form');
    const downloadBtn = document.getElementById('download-btn');
    const downloadMessage = document.getElementById('download-message');
    const homeworkNameInput = document.getElementById('homework-name');

    downloadForm.onsubmit = async (e) => {
      e.preventDefault();
      
      const homeworkName = homeworkNameInput.value.trim();
      if (!homeworkName) {
        showMessage(downloadMessage, '请输入作业名称。', 'red');
        return;
      }
      
      downloadBtn.disabled = true;
      downloadBtn.innerText = '正在生成...';
      showMessage(downloadMessage, '正在请求成绩总表, 请稍候...', 'blue');

      try {
        // 使用 /summary 路由, 并附带查询参数
        const url = \`/summary?homework=\${encodeURIComponent(homeworkName)}\`;
        const response = await fetch(url, { method: 'GET' });

        if (response.ok) {
          // 成功, 触发浏览器下载
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
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          
          showMessage(downloadMessage, '下载已开始!', 'green');

        } else {
          // 失败
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

    // --- 消息辅助函数 ---
    function showMessage(element, message, color) {
      element.innerHTML = \`<span class="text-\${color}-600">\${message}</span>\`;
    }
  </script>
</body>
</html>
  `;
}

