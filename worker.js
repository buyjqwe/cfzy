// --- Cloudflare Worker for Large File Analysis with OneDrive ---
// Architecture:
// 1. Client uploads a large ZIP file to the Worker.
// 2. The Worker streams the file directly to a Microsoft OneDrive account using the Graph API.
// 3. The Worker triggers a background analysis task.
// 4. The background task reads the ZIP from OneDrive, processes it in batches with Gemini.
// 5. The final report is saved back to OneDrive for the user to download.

// Import unzipit from a stable, version-pinned CDN URL as recommended.
// @ts-ignore
import { unzip } from 'https://unpkg.com/unzipit@1.4.3/dist/unzipit.module.js';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

// A simple in-memory cache for the Microsoft Graph API access token
let graphApiToken = {
    token: null,
    expiresAt: 0,
};

export default {
  async fetch(request, env, ctx) {
    if (request.method === 'OPTIONS') {
      // Handle CORS preflight requests
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    const url = new URL(request.url);

    if (url.pathname === '/') {
      // Serve the HTML UI
      return new Response(html, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }
    if (url.pathname === '/api/start-analysis' && request.method === 'POST') {
      // Handle file upload and start analysis
      return handleStartAnalysis(request, env, ctx);
    }
    if (url.pathname.startsWith('/results/') && request.method === 'GET') {
      // Handle polling for results
      return handleGetResult(url.pathname, env);
    }

    // Default 404 response
    return new Response('Not Found', { status: 404 });
  },
};

/**
 * Gets a valid Microsoft Graph API access token, caching it if possible.
 */
async function getGraphApiAccessToken(env) {
    if (graphApiToken.token && Date.now() < graphApiToken.expiresAt) {
        return graphApiToken.token;
    }

    const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = env;
    if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
        throw new Error("Microsoft API credentials are not configured in Worker secrets.");
    }

    const tokenUrl = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            client_id: MS_CLIENT_ID,
            client_secret: MS_CLIENT_SECRET,
            scope: 'https://graph.microsoft.com/.default',
            grant_type: 'client_credentials',
        }),
    });

    if (!response.ok) {
        const errorData = await response.text();
        console.error("Token fetch error:", errorData);
        throw new Error(`Failed to get Graph API token: ${response.status} ${response.statusText} - ${errorData}`);
    }

    const data = await response.json();
    graphApiToken = {
        token: data.access_token,
        expiresAt: Date.now() + (data.expires_in - 60) * 1000,
    };
    console.log("New access token acquired.");
    return graphApiToken.token;
}


/**
 * Handles the file upload FROM the client, streams it TO OneDrive, then starts analysis.
 */
async function handleStartAnalysis(request, env, ctx) {
    try {
        const formData = await request.formData();
        const userPrompt = formData.get('userPrompt');
        const file = formData.get('file');

        if (!file || !userPrompt || !(file instanceof File)) {
            return new Response(JSON.stringify({ error: '请求无效，缺少文件或指令。' }), { 
                status: 400, 
                headers: { ...corsHeaders, 'Content-Type': 'application/json' } 
            });
        }
        
        const accessToken = await getGraphApiAccessToken(env);
        const userId = env.MS_USER_ID || 'me';
        
        const fileName = `uploads/${Date.now()}-${file.name}`;
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const resultKey = `results/${timestamp}_${file.name.replace(/\.[^/.]+$/, '')}_analysis.txt`;

        const sessionUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${fileName}:/createUploadSession`;
        const sessionResponse = await fetch(sessionUrl, {
            method: 'POST',
            headers: { 
                'Authorization': `Bearer ${accessToken}`, 
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
        });

        if (!sessionResponse.ok) {
            const sessionError = await sessionResponse.text();
            throw new Error(`无法创建 OneDrive 上传会话: ${sessionResponse.statusText}. ${sessionError}`);
        }
        const { uploadUrl } = await sessionResponse.json();

        if (!uploadUrl) {
            throw new Error("OneDrive did not return a valid upload URL.");
        }

        const uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: { 'Content-Length': file.size.toString() },
            body: file.stream(),
        });

        if (!uploadResponse.ok) {
            const uploadError = await uploadResponse.text();
            throw new Error(`上传文件到 OneDrive 失败: ${uploadResponse.statusText}. ${uploadError}`);
        }

        const uploadedFile = await uploadResponse.json();
        let itemId = uploadedFile.id;

        if (!itemId) {
             console.warn("Upload response did not contain an ID. Querying by path as a fallback.");
             const itemByPathUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${fileName}:/`;
             const itemResponse = await fetch(itemByPathUrl, {
                 headers: { 'Authorization': `Bearer ${accessToken}` }
             });
             if (!itemResponse.ok) {
                 throw new Error("上传成功，但无法通过路径获取文件ID。");
             }
             const itemData = await itemResponse.json();
             itemId = itemData.id;
        }

        if (!itemId) {
            throw new Error("无法最终确定上传文件的 OneDrive Item ID。");
        }

        console.log(`File uploaded. Item ID: ${itemId}`);

        ctx.waitUntil(performAnalysis(itemId, resultKey, userPrompt, env));

        return new Response(JSON.stringify({ 
            message: '文件上传成功，分析已在后台启动。', 
            resultKey: resultKey,
        }), {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
        });

    } catch(e) {
        console.error("handleStartAnalysis Error:", e);
        return new Response(JSON.stringify({ error: `启动分析失败: ${e.message}` }), { 
            status: 500, 
            headers: { 'Content-Type': 'application/json', ...corsHeaders } 
        });
    }
}

/**
 * Serves the result file from OneDrive when the client polls for it.
 */
async function handleGetResult(pathname, env) {
    const resultKey = pathname.substring(1);
    try {
        const accessToken = await getGraphApiAccessToken(env);
        const userId = env.MS_USER_ID || 'me';
        const downloadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
        
        const response = await fetch(downloadUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` },
        });

        if (response.status === 404) {
             return new Response('分析结果尚未准备好。', { status: 404, headers: corsHeaders });
        }
        
        if (!response.ok) {
             return new Response(`获取结果时出错: ${response.statusText}`, { status: response.status, headers: corsHeaders });
        }
        
        const customHeaders = new Headers(response.headers);
        customHeaders.set('Content-Type', 'text/plain; charset=utf-8');
        
        return new Response(response.body, { headers: customHeaders });

    } catch (e) {
        console.error("handleGetResult Error:", e);
        return new Response('获取结果时发生内部错误。', { status: 500, headers: corsHeaders });
    }
}


/**
 * The core analysis function that runs in the background.
 */
async function performAnalysis(driveItemId, resultKey, userPrompt, env) {
  console.log(`Starting background analysis for item: ${driveItemId}`);
  try {
    const accessToken = await getGraphApiAccessToken(env);
    const userId = env.MS_USER_ID || 'me';
    const fileUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${driveItemId}/content`;

    const response = await fetch(fileUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!response.ok) {
        throw new Error(`无法从 OneDrive 下载文件: ${response.status} ${response.statusText}`);
    }

    const { entries } = await unzip(response.body);

    const fileTypes = {
      '图像 (Image)': ['.png', '.jpg', '.jpeg', '.webp', '.gif'],
      '核心文本/代码 (Core Text/Code)': ['.html', '.css', '.md', '.py', '.cs', '.json', '.js', '.txt', '.ts', '.tsx', '.jsx', '.java', '.cpp', '.c', '.h', '.sql', '.xml', '.yaml', '.yml'],
      '其他文本 (Other Text)': ['.csv', '.log', '.rtf', '.tex', '.bib']
    };
    let finalReport = `Gemini 分析报告\n==================\n\n用户指令: ${userPrompt}\n\n`;
    let fileCount = 0;
    const BATCH_SIZE_LIMIT_BYTES = 3.5 * 1024 * 1024;
    let currentBatchParts = [];
    let currentBatchSizeBytes = 0;

    for (const [filePath, entry] of Object.entries(entries)) {
      if (entry.isDirectory) continue;
      
      fileCount++;
      const fileData = new Uint8Array(await entry.arrayBuffer());
      if (fileData.length === 0) continue;

      const extension = `.${filePath.split('.').pop()?.toLowerCase()}` || '';
      const contentHeader = `\n\n--- 文件: ${filePath} ---\n`;
      let partSegments = [];

      if (fileTypes['图像 (Image)'].includes(extension)) {
          const mimeType = getMimeType(extension);
          // Fixed: Move the base64 assignment before the size check
          const base64 = arrayBufferToBase64(fileData.buffer);
          if (base64.length > 4 * 1024 * 1024) {
              partSegments = [{ text: contentHeader + "[图像文件过大，无法分析。]" }];
          } else {
              partSegments = [{ text: contentHeader }, { inlineData: { mimeType, data: base64 } }];
          }
      } else if (Object.values(fileTypes).flat().includes(extension)) {
           try {
              let textContent = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(fileData);
              textContent = sanitizeText(textContent);
              if (textContent.length > 100000) {
                  textContent = textContent.substring(0, 100000) + "\n... [内容已被截断] ...";
              }
              partSegments = [{ text: contentHeader + textContent }];
           } catch (e) {
              partSegments = [{ text: contentHeader + `[文件无法以UTF-8解码，已跳过。]` }];
           }
      } else {
          partSegments = [{ text: contentHeader + "[不支持分析的文件类型，已跳过。]" }];
      }

      const partSize = JSON.stringify(partSegments).length;
      if (currentBatchSizeBytes + partSize > BATCH_SIZE_LIMIT_BYTES && currentBatchParts.length > 0) {
        finalReport += await processBatch(currentBatchParts, userPrompt, env);
        currentBatchParts = [];
        currentBatchSizeBytes = 0;
      }
      
      currentBatchParts.push(...partSegments);
      currentBatchSizeBytes += partSize;
    }

    if (currentBatchParts.length > 0) {
      finalReport += await processBatch(currentBatchParts, userPrompt, env);
    }
    
    finalReport += `\n\n==================\n分析完成。共处理 ${fileCount} 个文件。`;
    
    const reportUploadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
    await fetch(reportUploadUrl, {
        method: 'PUT',
        headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'text/plain; charset=utf-8'
        },
        body: finalReport,
    });

  } catch (e) {
    console.error(`后台分析失败 for item ${driveItemId}:`, e);
    try {
        const accessToken = await getGraphApiAccessToken(env);
        if (accessToken) {
            const userId = env.MS_USER_ID || 'me';
            const errorReportUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
            const errorReportContent = `分析失败: ${e.message}\n\n详细错误信息:\n${e.stack}`;
            await fetch(errorReportUrl, {
                method: 'PUT',
                headers: { 
                    'Authorization': `Bearer ${accessToken}`, 
                    'Content-Type': 'text/plain; charset=utf-8'
                },
                body: errorReportContent,
            });
        }
    } catch (tokenError) {
        console.error("Could not save error report due to token error:", tokenError);
    }
  }
}

/**
 * Sends a batch of file parts to the Gemini API for analysis.
 */
async function processBatch(parts, userPrompt, env) {
    const GOOGLE_API_KEY = env.GOOGLE_API_KEY;
    if (!GOOGLE_API_KEY) {
        throw new Error("未配置 'GOOGLE_API_KEY' 环境变量。");
    }

    const model = 'gemini-2.5-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    
    const payload = { 
        contents: [{ role: "user", parts: [{ text: userPrompt }, ...parts] }],
    };

    try {
        const apiResponse = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        if (!apiResponse.ok) {
            const errorBody = await apiResponse.text();
            let errorMessage = `API Error ${apiResponse.status}: ${apiResponse.statusText}`;
            try {
                const errorJson = JSON.parse(errorBody);
                if (errorJson.error?.message) {
                      errorMessage += `. Details: ${errorJson.error.message}`;
                }
            } catch (e) {
                errorMessage += `. Raw response: ${errorBody.substring(0, 200)}...`;
            }
            return `\n[分析批次失败: ${errorMessage}]\n`;
        }

        const resultJson = await apiResponse.json();
        
        if (!resultJson.candidates || resultJson.candidates.length === 0) {
            const blockReason = resultJson.promptFeedback?.blockReason || 'No candidates returned';
            return `\n[模型未返回有效内容 (可能因安全原因被阻止: ${blockReason})]\n`;
        }
        
        const generatedText = resultJson.candidates[0]?.content?.parts?.[0]?.text;
        if (typeof generatedText !== 'string' || generatedText.trim() === '') {
            return "\n[模型返回了结果，但未包含有效文本内容]\n";
        }

        return generatedText;

    } catch (e) {
       console.error("Network error calling Gemini API:", e);
       return `\n[调用 Gemini API 时发生网络错误: ${e.message}]\n`;
    }
}

// --- Helper Functions ---
function getMimeType(extension) {
    const mimeTypes = { 
        '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', 
        '.gif': 'image/gif', '.webp': 'image/webp', '.bmp': 'image/bmp',
        '.svg': 'image/svg+xml', '.tiff': 'image/tiff', '.ico': 'image/x-icon'
    };
    return mimeTypes[extension.toLowerCase()] || 'application/octet-stream';
}

function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

function sanitizeText(text) {
  if (typeof text !== 'string') return '';
  return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, ''); 
}


// --- HTML User Interface ---
const html = 
'<!DOCTYPE html>' +
'<html lang="zh-CN">' +
'<head>' +
'    <meta charset="UTF-8">' +
'    <meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'    <title>Gemini 大型文件分析器 (OneDrive版)</title>' +
'    <script src="https://cdn.tailwindcss.com"></script>' +
'    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">' +
'    <style>' +
'        body { font-family: \'Inter\', sans-serif; background-color: #f3f4f6; }' +
'        .container { max-width: 800px; }' +
'        #result-box, #status-box {' +
'            white-space: pre-wrap; word-wrap: break-word; background-color: #1e293b;' +
'            color: #e2e8f0; border-radius: 0.5rem; padding: 1.5rem; margin-top: 1.5rem;' +
'            line-height: 1.75; font-family: \'Courier New\', Courier, monospace;' +
'            overflow-x: auto;' +
'        }' +
'        .loader {' +
'            display: inline-block;' +
'            border: 4px solid #f3f3f3; border-top: 4px solid #3b82f6; border-radius: 50%;' +
'            width: 20px; height: 20px; animation: spin 1s linear infinite; margin-right: 10px; vertical-align: middle;' +
'        }' +
'        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
'    </style>' +
'</head>' +
'<body class="antialiased text-gray-800">' +
'    <div class="container mx-auto p-4 sm:p-6 lg:p-8">' +
'        <header class="text-center mb-8">' +
'            <h1 class="text-3xl sm:text-4xl font-bold text-gray-900">Gemini 大型文件分析器</h1>' +
'            <p class="mt-2 text-lg text-gray-600">由 OneDrive & Gemini 强力驱动</p>' +
'        </header>' +
'        <form id="upload-form" class="bg-white p-6 rounded-lg shadow-md border border-gray-200">' +
'            <div class="mb-5">' +
'                <label for="user_prompt" class="block mb-2 text-md font-medium text-gray-700">你的分析指令:</label>' +
'                <textarea id="user_prompt" name="userPrompt" rows="4" class="w-full p-3 border border-gray-300 rounded-md" placeholder="例如：请总结每个代码文件的主要功能和潜在问题。" required></textarea>' +
'            </div>' +
'            <div class="mb-5">' +
'                <label for="zip-upload" class="block mb-2 text-md font-medium text-gray-700">选择一个 ZIP 压缩包上传:</label>' +
'                <input type="file" name="file" id="zip-upload" accept=".zip" class="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:bg-blue-600 file:text-white hover:file:bg-blue-700" required/>' +
'            </div>' +
'            <button type="submit" id="submit-btn" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-4 rounded-md disabled:opacity-50">' +
'                开始分析' +
'            </button>' +
'        </form>' +
'        <div id="status-container" class="mt-6"></div>' +
'    </div>' +
'    <script>' +
'        const form = document.getElementById(\'upload-form\');' +
'        const submitBtn = document.getElementById(\'submit-btn\');' +
'        const zipUpload = document.getElementById(\'zip-upload\');' +
'        const statusContainer = document.getElementById(\'status-container\');' +
'        form.addEventListener(\'submit\', async (e) => {' +
'            e.preventDefault();' +
'            const file = zipUpload.files[0];' +
'            if (!file) { updateStatus("错误：请选择一个 ZIP 文件。", "error"); return; }' +
'            setLoadingState(true, "正在上传文件并启动分析... 这可能需要几分钟，请勿关闭页面。");' +
'            try {' +
'                const formData = new FormData(form);' +
'                const analysisResponse = await fetch("/api/start-analysis", { method: "POST", body: formData });' +
'                const responseText = await analysisResponse.text();' +
'                if (!analysisResponse.ok) {' +
'                    let msg = "启动分析任务失败。";' +
'                    try { msg = JSON.parse(responseText).error || msg; } catch (err) { msg = responseText || msg; }' +
'                    throw new Error(msg);' +
'                }' +
'                const analysisData = JSON.parse(responseText);' +
'                updateStatus("文件上传成功！分析任务已在后台运行。正在等待最终报告...", "loading");' +
'                pollForResult(analysisData.resultKey);' +
'            } catch (error) {' +
'                console.error("Submit Error:", error);' +
'                updateStatus("发生错误：" + error.message, "error");' +
'                setLoadingState(false);' +
'            }' +
'        });' +
'        function pollForResult(resultKey) {' +
'            const pollInterval = 15000;' +
'            const maxAttempts = 120;' +
'            let attempts = 0;' +
'            const intervalId = setInterval(async () => {' +
'                if (attempts++ > maxAttempts) {' +
'                    clearInterval(intervalId);' +
'                    updateStatus("分析超时（30分钟）。", "error");' +
'                    setLoadingState(false);' +
'                    return;' +
'                }' +
'                try {' +
'                    const resultResponse = await fetch("/" + resultKey);' +
'                    if (resultResponse.status === 200) {' +
'                        clearInterval(intervalId);' +
'                        const reportText = await resultResponse.text();' +
'                        updateStatus("分析完成！", "success");' +
'                        displayResult(reportText, resultKey);' +
'                        setLoadingState(false);' +
'                    } else if (resultResponse.status !== 404) {' +
'                        clearInterval(intervalId);' +
'                        const errorText = await resultResponse.text();' +
'                        throw new Error(`获取结果时发生服务器错误: ${resultResponse.status} ${errorText}`);' +
'                    }' +
'                } catch (error) {' +
'                    clearInterval(intervalId);' +
'                    console.error("Polling Error:", error);' +
'                    updateStatus("获取结果失败：" + error.message, "error");' +
'                    setLoadingState(false);' +
'                }' +
'            }, pollInterval);' +
'        }' +
'        function setLoadingState(isLoading, message = "") {' +
'            submitBtn.disabled = isLoading;' +
'            submitBtn.textContent = isLoading ? "处理中..." : "开始分析";' +
'            if (isLoading) updateStatus(message, "loading");' +
'        }' +
'        function updateStatus(message, type) {' +
'            let icon = "";' +
'            if (type === "loading") icon = \'<div class="loader"></div>\';' +
'            let color = type === "error" ? "text-red-500" : (type === "success" ? "text-green-500" : "");' +
'            statusContainer.innerHTML = \'<div id="status-box"><p class="\' + color + \' font-semibold">\' + icon + \'<span>\' + message + \'</span></p></div>\';' +
'        }' +
'        function displayResult(reportText, resultKey) {' +
'            const sanitizedText = reportText.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");' +
'            const resultHtml = \'<div id="result-box"><a href="/\' + resultKey + \'" target="_blank" class="float-right bg-gray-600 text-white py-1 px-3 rounded-md text-sm">打开报告</a><pre>\' + sanitizedText + \'</pre></div>\';' +
'            statusContainer.innerHTML += resultHtml;' +
'        }' +
'    </script>' +
'</body>' +
'</html>';
}



