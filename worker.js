// --- Cloudflare Worker for Large File Analysis with OneDrive ---
// FINAL, INTELLIGENT VERSION:
// 1. Client UI accepts either a folder or a single ZIP file.
// 2. Client-side JS intelligently zips the folder in-browser or uses the provided ZIP directly.
// 3. The single ZIP is streamed to OneDrive, bypassing size limits.
// 4. The background worker performs nested unzipping, handling ZIPs within the main ZIP.

// @ts-ignore
import { unzip } from './unzipit.module.js';

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
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    const url = new URL(request.url);

    if (url.pathname === '/') {
      return new Response(html, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }
    if (url.pathname === '/api/start-analysis' && request.method === 'POST') {
      return handleStartAnalysis(request, env, ctx);
    }
    if (url.pathname.startsWith('/results/') && request.method === 'GET') {
      return handleGetResult(url.pathname, env);
    }

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
        throw new Error(`Failed to get Graph API token: ${errorData}`);
    }

    const data = await response.json();
    graphApiToken = {
        token: data.access_token,
        expiresAt: Date.now() + (data.expires_in - 60) * 1000,
    };
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

        if (!file || !userPrompt) {
            return new Response(JSON.stringify({ error: '请求无效，缺少文件或指令。' }), { status: 400, headers: corsHeaders });
        }
        
        const accessToken = await getGraphApiAccessToken(env);
        const userId = env.MS_USER_ID || 'me';
        
        const fileName = `uploads/${Date.now()}-${file.name}`;
        const resultKey = `results/${Date.now()}-${file.name.replace(/\.zip$/i, '.txt')}`;

        const sessionUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${fileName}:/createUploadSession`;
        const sessionResponse = await fetch(sessionUrl, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
        });

        if (!sessionResponse.ok) throw new Error('无法创建 OneDrive 上传会话。');
        const { uploadUrl } = await sessionResponse.json();

        const uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: { 'Content-Length': file.size },
            body: file.stream(),
        });

        if (!uploadResponse.ok) throw new Error('上传文件到 OneDrive 失败。');

        const uploadedFile = await uploadResponse.json();

        ctx.waitUntil(performAnalysis(uploadedFile.id, resultKey, userPrompt, env));

        return new Response(JSON.stringify({ message: '分析已开始，请稍后查看结果。', resultKey: resultKey }), {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
        });

    } catch(e) {
        console.error("Analysis Start/Upload Error:", e);
        return new Response(JSON.stringify({ error: `启动分析失败: ${e.message}` }), { status: 500, headers: corsHeaders });
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

        if (!response.ok) {
             return new Response('分析结果尚未准备好，或不存在。', { status: 404 });
        }
        
        return new Response(response.body, { headers: response.headers });

    } catch (e) {
        console.error("Get Result Error:", e);
        return new Response('获取结果失败。', { status: 500 });
    }
}


/**
 * The core analysis function that runs in the background.
 */
async function performAnalysis(driveItemId, resultKey, userPrompt, env) {
  try {
    const accessToken = await getGraphApiAccessToken(env);
    const userId = env.MS_USER_ID || 'me';
    const fileUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${driveItemId}/content`;

    const response = await fetch(fileUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!response.ok) throw new Error(`无法从 OneDrive 下载文件: ${driveItemId}`);
    
    const { entries: topLevelEntries } = await unzip(response.body);

    const fileTypes = {
      '图像 (Image)': ['.png', '.jpg', '.jpeg', '.webp', '.gif'],
      '核心文本/代码 (Core Text/Code)': ['.html', '.css', '.md', '.py', '.cs', '.json', '.js', '.txt'],
    };
    let finalReport = `Gemini 分析报告\n==================\n\n用户指令: ${userPrompt}\n\n`;
    let fileCount = 0;
    const BATCH_SIZE_LIMIT_BYTES = 4 * 1024 * 1024;
    let currentBatchParts = [];
    let currentBatchSizeBytes = 0;

    // Helper function to process a file entry and add it to the batch
    const processFileEntry = async (entry, filePath) => {
        if (entry.isDirectory) return;
        fileCount++;
        const fileData = new Uint8Array(await entry.arrayBuffer());
        if (fileData.length === 0) return;

        const extension = `.${filePath.split('.').pop()?.toLowerCase()}`;
        const contentHeader = `\n\n--- 文件: ${filePath} ---\n`;
        let partSegments = [];

        if (fileTypes['图像 (Image)'].includes(extension)) {
            const mimeType = getMimeType(extension);
            const base64 = arrayBufferToBase64(fileData.buffer);
            partSegments = [{ text: contentHeader }, { inlineData: { mimeType, data: base64 } }];
        } else if (fileTypes['核心文本/代码 (Core Text/Code)'].includes(extension)) {
            try {
                const textContent = new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(fileData);
                partSegments = [{ text: contentHeader + sanitizeText(textContent) }];
            } catch (e) {
                partSegments = [{ text: contentHeader + "[文件内容无法被识别为UTF-8编码，已跳过。]" }];
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
    };

    // Process all entries from the top-level ZIP
    for (const [filePath, entry] of Object.entries(topLevelEntries)) {
        // If an entry is a nested ZIP file, unzip it and process its contents
        if (filePath.toLowerCase().endsWith('.zip')) {
            console.log(`Found nested ZIP: ${filePath}. Unzipping...`);
            const nestedZipData = await entry.arrayBuffer();
            const { entries: nestedEntries } = await unzip(new Uint8Array(nestedZipData));
            for (const [nestedFilePath, nestedEntry] of Object.entries(nestedEntries)) {
                const fullPath = `${filePath}/${nestedFilePath}`; // Create a full path for context
                await processFileEntry(nestedEntry, fullPath);
            }
        } else {
            // Otherwise, process it as a regular file
            await processFileEntry(entry, filePath);
        }
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
    console.error("后台分析失败:", e);
    const accessToken = await getGraphApiAccessToken(env).catch(() => null);
    if (accessToken) {
        const userId = env.MS_USER_ID || 'me';
        const errorReportUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
        await fetch(errorReportUrl, {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'text/plain' },
            body: `分析失败: ${e.message}\n\n请检查您的文件或联系支持。`
        });
    }
  }
}

async function processBatch(parts, userPrompt, env) {
    const GOOGLE_API_KEY = env.GOOGLE_API_KEY;
    if (!GOOGLE_API_KEY) throw new Error("未配置 'GOOGLE_API_KEY'。");
    const model = 'gemini-2.5-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    const payload = { contents: [{ parts: [{text: userPrompt}, ...parts] }] };
    const apiResponse = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!apiResponse.ok) {
        const errorBody = await apiResponse.text();
        console.error("Gemini API Error:", errorBody);
        return `\n[一个分析批次失败: ${apiResponse.statusText}]\n`;
    }
    const resultJson = await apiResponse.json();
    return resultJson.candidates?.[0]?.content?.parts?.[0]?.text || "\n[模型未返回有效内容]\n";
}

// --- Helper Functions ---
function getMimeType(extension) {
    const mimeTypes = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp' };
    return mimeTypes[extension] || 'application/octet-stream';
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
'    <title>Gemini 智能作业分析器</title>' +
'    <script src="https://cdn.tailwindcss.com"></script>' +
'    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>' +
'    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">' +
'    <style>' +
'        body { font-family: \'Inter\', sans-serif; background-color: #f3f4f6; }' +
'        .container { max-width: 800px; }' +
'        #result-box, #status-box {' +
'            white-space: pre-wrap; word-wrap: break-word; background-color: #1e293b;' +
'            color: #e2e8f0; border-radius: 0.5rem; padding: 1.5rem; margin-top: 1.5rem;' +
'            line-height: 1.75; font-family: \'Courier New\', Courier, monospace;' +
'        }' +
'        .loader {' +
'            border: 4px solid #f3f3f3; border-top: 4px solid #3b82f6; border-radius: 50%;' +
'            width: 20px; height: 20px; animation: spin 1s linear infinite; margin-right: 10px; vertical-align: middle; display: inline-block;' +
'        }' +
'        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
'    </style>' +
'</head>' +
'<body class="antialiased text-gray-800">' +
'    <div class="container mx-auto p-4 sm:p-6 lg:p-8">' +
'        <header class="text-center mb-8">' +
'            <h1 class="text-3xl sm:text-4xl font-bold text-gray-900">Gemini 智能作业分析器</h1>' +
'            <p class="mt-2 text-lg text-gray-600">2510222131</p>' +
'        </header>' +
'        <form id="upload-form" class="bg-white p-6 rounded-lg shadow-md border border-gray-200">' +
'            <div class="mb-5">' +
'                <label for="user_prompt" class="block mb-2 text-md font-medium text-gray-700">你的分析指令:</label>' +
'                <textarea id="user_prompt" name="userPrompt" rows="4" class="w-full p-3 border border-gray-300 rounded-md" placeholder="例如：请分别总结每个学生的作业情况，并进行打分。" required></textarea>' +
'            </div>' +
'            <div class="mb-5">' +
'                 <label for="file-upload" class="block mb-2 text-md font-medium text-gray-700">选择作业文件夹 或 作业ZIP包:</label>' +
'                <input type="file" name="files" id="file-upload" webkitdirectory directory multiple class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:bg-blue-600 file:text-white hover:file:bg-blue-700" required/>' +
'            </div>' +
'            <button type="submit" id="submit-btn" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-md disabled:opacity-50">开始分析</button>' +
'        </form>' +
'        <div id="status-container" class="mt-6"></div>' +
'    </div>' +
'    <script>' +
'        const form = document.getElementById(\'upload-form\');' +
'        const submitBtn = document.getElementById(\'submit-btn\');' +
'        const fileUpload = document.getElementById(\'file-upload\');' +
'        const statusContainer = document.getElementById(\'status-container\');' +
'        form.addEventListener(\'submit\', async (e) => {' +
'            e.preventDefault();' +
'            const files = fileUpload.files;' +
'            if (files.length === 0) { updateStatus("错误：请选择一个文件夹或ZIP文件。", "error"); return; }' +
'            setLoadingState(true, "正在准备文件...");' +
'            try {' +
'                let fileToUpload;' +
'                let fileName;' +
'                if (files.length > 1 || (files[0] && files[0].webkitRelativePath)) {' +
'                    updateStatus("正在浏览器中打包文件夹...", "loading");' +
'                    const zip = new JSZip();' +
'                    let folderName = "";' +
'                    for (const file of files) {' +
'                        zip.file(file.webkitRelativePath, file);' +
'                        if (!folderName) { folderName = file.webkitRelativePath.split("/")[0]; }' +
'                    }' +
'                    fileToUpload = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 1 } });' +
'                    fileName = folderName + ".zip";' +
'                } else if (files.length === 1) {' +
'                    const singleFile = files[0];' +
'                    if (!singleFile.name.toLowerCase().endsWith(".zip")) {' +
'                         updateStatus("错误：如果您只选择一个文件，它必须是 .zip 格式。", "error");' +
'                         setLoadingState(false);' +
'                         return;' +
'                    }' +
'                    fileToUpload = singleFile;' +
'                    fileName = singleFile.name;' +
'                }' +
'                await uploadAndStartAnalysis(fileToUpload, fileName);' +
'            } catch (error) {' +
'                console.error("Submit Error:", error);' +
'                updateStatus("发生错误：" + error.message, "error");' +
'                setLoadingState(false);' +
'            }' +
'        });' +
'        async function uploadAndStartAnalysis(fileBlob, fileName) {' +
'            updateStatus("文件准备就绪，正在上传到OneDrive...", "loading");' +
'            const formData = new FormData();' +
'            formData.append("file", fileBlob, fileName);' +
'            formData.append("userPrompt", document.getElementById("user_prompt").value);' +
'            const analysisResponse = await fetch("/api/start-analysis", { method: "POST", body: formData });' +
'            const responseText = await analysisResponse.text();' +
'            if (!analysisResponse.ok) {' +
'                let msg = "启动分析任务失败。";' +
'                try { msg = JSON.parse(responseText).error || msg; } catch (err) { msg = responseText || msg; }' +
'                throw new Error(msg);' +
'            }' +
'            const analysisData = JSON.parse(responseText);' +
'            updateStatus("文件上传成功！分析任务已在后台运行...", "loading");' +
'            pollForResult(analysisData.resultKey);' +
'        }' +
'        function pollForResult(resultKey) {' +
'            const pollInterval = 10000;' +
'            const maxAttempts = 180;' +
'            let attempts = 0;' +
'            const intervalId = setInterval(async () => {' +
'                if (attempts++ > maxAttempts) {' +
'                    clearInterval(intervalId);' +
'                    updateStatus("分析超时。", "error");' +
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
'            let icon = type === "loading" ? \'<div class="loader"></div>\' : "";' +
'            let color = type === "error" ? "text-red-500" : (type === "success" ? "text-green-500" : "");' +
'            statusContainer.innerHTML = \'<div id="status-box"><p class="\' + color + \' font-semibold">\' + icon + \'<span>\' + message + \'</span></p></div>\';' +
'        }' +
'        function displayResult(reportText, resultKey) {' +
'            const sanitizedText = reportText.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");' +
'            const resultHtml = \'<div id="result-box"><a href="/\' + resultKey + \'" download class="float-right bg-gray-600 text-white py-1 px-3 rounded-md text-sm">下载报告</a><pre>\' + sanitizedText + \'</pre></div>\';' +
'            statusContainer.innerHTML += resultHtml;' +
'        }' +
'    </script>' +
'</body>' +
'</html>';

