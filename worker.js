// --- Cloudflare Worker for Large File Analysis with OneDrive ---
// Architecture:
// 1. Client uploads a large ZIP file to the Worker.
// 2. The Worker streams the file directly to a Microsoft OneDrive account using the Graph API.
// 3. The Worker triggers a background analysis task.
// 4. The background task reads the ZIP from OneDrive, processes it in batches with Gemini.
// 5. The final report is saved back to OneDrive for the user to download.

// Import unzipit from a more stable ESM source
// @ts-ignore
import { unzip } from 'https://cdn.jsdelivr.net/npm/unzipit@0.6.3/dist/unzipit.esm.js';

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
 * This uses the Client Credentials Grant Flow, suitable for server-to-server interaction.
 */
async function getGraphApiAccessToken(env) {
    // Check if the current token is still valid
    if (graphApiToken.token && Date.now() < graphApiToken.expiresAt) {
        return graphApiToken.token;
    }

    const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = env;
    if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
        throw new Error("Microsoft API credentials are not configured in Worker secrets (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET).");
    }

    const tokenUrl = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', MS_CLIENT_ID);
    params.append('client_secret', MS_CLIENT_SECRET);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params,
    });

    if (!response.ok) {
        const errorData = await response.text();
        console.error("Token fetch error:", errorData);
        throw new Error(`Failed to get Graph API token: ${response.status} ${response.statusText} - ${errorData}`);
    }

    const data = await response.json();
    graphApiToken = {
        token: data.access_token,
        // Set expiry to 1 minute before the actual expiry time for safety
        expiresAt: Date.now() + (data.expires_in - 60) * 1000,
    };
    console.log("New access token acquired, expires in:", data.expires_in, "seconds");
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
            return new Response(JSON.stringify({ error: '请求无效，缺少文件或指令，或文件格式不正确。' }), { 
                status: 400, 
                headers: { ...corsHeaders, 'Content-Type': 'application/json' } 
            });
        }
        
        const accessToken = await getGraphApiAccessToken(env);
        const userId = env.MS_USER_ID || 'me'; // Use 'me' or a specific user ID/email
        
        const fileName = `uploads/${Date.now()}-${file.name}`;
        // Use a more descriptive name for the result file
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const resultKey = `results/${timestamp}_${file.name.replace(/\.[^/.]+$/, '')}_analysis.txt`;

        // Use the large file upload mechanism for OneDrive (Create Upload Session)
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
            console.error("Create upload session error:", sessionError);
            throw new Error(`无法创建 OneDrive 上传会话: ${sessionResponse.status} ${sessionResponse.statusText}. ${sessionError}`);
        }
        const sessionData = await sessionResponse.json();
        const { uploadUrl } = sessionData;

        if (!uploadUrl) {
            throw new Error("OneDrive 上传会话创建成功，但未返回有效的上传 URL。");
        }

        // Stream the file upload directly to the OneDrive upload URL
        // The uploadUrl expects specific headers and range for large files, but for simplicity here,
        // we assume the initial session handles the first chunk or the entire file if it's small enough for a single PUT.
        // For truly large files, multi-chunk upload logic would be needed here.
        const uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: { 
                'Content-Length': file.size.toString() // Ensure it's a string
            },
            body: file.stream(),
        });

        if (!uploadResponse.ok) {
            const uploadError = await uploadResponse.text();
            console.error("File upload error:", uploadError);
            throw new Error(`上传文件到 OneDrive 失败: ${uploadResponse.status} ${uploadResponse.statusText}. ${uploadError}`);
        }

        const uploadedFile = await uploadResponse.json();
        const itemId = uploadedFile.id;

        if (!itemId) {
             // If itemId is not available in the upload completion response, we might need to query for it by path.
             // However, the completion response *should* contain the item details including the ID.
             console.warn("Upload completion response did not contain 'id'. Attempting to find file by path...");
             // This step is complex and might require another API call based on the fileName.
             // For now, let's assume the ID is present or fail gracefully.
             // A more robust solution would be to list the 'uploads' folder and find the most recent file matching the pattern.
             // As a fallback, we can try to get the item by its path.
             const itemByPathUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${fileName}:/`;
             const itemResponse = await fetch(itemByPathUrl, {
                 headers: { 'Authorization': `Bearer ${accessToken}` }
             });
             if (!itemResponse.ok) {
                 console.error("Could not retrieve item by path after upload.");
                 throw new Error("上传成功，但无法获取文件ID以进行后续分析。");
             }
             const itemData = await itemResponse.json();
             itemId = itemData.id;
        }

        if (!itemId) {
            throw new Error("无法确定上传文件的 OneDrive Item ID。");
        }

        console.log(`File uploaded successfully. Item ID: ${itemId}, Path: ${fileName}`);

        // Start the long-running analysis in the background.
        ctx.waitUntil(performAnalysis(itemId, resultKey, userPrompt, env));

        return new Response(JSON.stringify({ 
            message: '文件上传成功，分析已在后台启动。请稍后查看结果。', 
            resultKey: resultKey,
            fileId: itemId // For potential future use or debugging
        }), {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
        });

    } catch(e) {
        console.error("Analysis Start/Upload Error:", e);
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
    const resultKey = pathname.substring(1); // remove leading '/'
    console.log("Polling for result at path:", resultKey);
    try {
        const accessToken = await getGraphApiAccessToken(env);
        const userId = env.MS_USER_ID || 'me';
        const downloadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
        
        const response = await fetch(downloadUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` },
        });

        if (response.status === 404) {
            // File not found, likely still being processed
             console.log("Result file not found yet (404):", resultKey);
             return new Response('分析结果尚未准备好，请稍后再试。', { status: 404, headers: corsHeaders });
        }
        
        if (!response.ok) {
             console.error("Error fetching result:", response.status, response.statusText);
             return new Response(`获取结果失败: ${response.status} ${response.statusText}`, { status: 500, headers: corsHeaders });
        }
        
        console.log("Result file found and being served:", resultKey);
        // Stream the response back to the client
        // Clone the response to potentially read headers
        const responseClone = response.clone();
        // Ensure the content type is text/plain for the browser to handle it correctly as a text file
        const customHeaders = new Headers(responseClone.headers);
        customHeaders.set('Content-Type', 'text/plain; charset=utf-8');
        // Optionally, force download: customHeaders.set('Content-Disposition', 'attachment; filename="analysis_report.txt"');
        
        return new Response(responseClone.body, { headers: customHeaders });

    } catch (e) {
        console.error("Get Result Error:", e);
        return new Response('获取结果时发生内部错误。', { status: 500, headers: corsHeaders });
    }
}


/**
 * The core analysis function that runs in the background.
 */
async function performAnalysis(driveItemId, resultKey, userPrompt, env) {
  console.log("Starting background analysis for item:", driveItemId);
  try {
    const accessToken = await getGraphApiAccessToken(env);
    const userId = env.MS_USER_ID || 'me';
    // Construct the URL to download the content of the uploaded ZIP file using its itemId
    const fileUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${driveItemId}/content`;

    const response = await fetch(fileUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!response.ok) {
        console.error(`Failed to download ZIP file content for item ${driveItemId}:`, response.status, response.statusText);
        // Attempt to get item metadata to confirm existence
        const metadataUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${driveItemId}`;
        const metaResponse = await fetch(metadataUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        if (!metaResponse.ok) {
            console.error("Could not get metadata for item, it might not exist:", driveItemId);
             throw new Error(`无法从 OneDrive 获取文件内容或元数据: ${driveItemId} (${response.status} / ${metaResponse.status})`);
        }
        const metadata = await metaResponse.json();
        console.log("Item metadata found:", metadata);
        // If metadata exists but content doesn't, it's an unusual state. Re-throw original error.
        throw new Error(`无法从 OneDrive 下载文件内容: ${driveItemId}, status: ${response.status} ${response.statusText}`);
    }

    console.log("Successfully downloaded ZIP content, starting to unzip...");
    // Use the unzipit library to decompress the ZIP stream
    const { entries } = await unzip(response.body);

    console.log("ZIP unzipped, found", Object.keys(entries).length, "entries. Starting analysis...");
    // Define file types for processing
    const fileTypes = {
      '图像 (Image)': ['.png', '.jpg', '.jpeg', '.webp', '.gif'],
      '核心文本/代码 (Core Text/Code)': ['.html', '.css', '.md', '.py', '.cs', '.json', '.js', '.txt', '.ts', '.tsx', '.jsx', '.java', '.cpp', '.c', '.h', '.sql', '.xml', '.yaml', '.yml'],
      '其他文本 (Other Text)': ['.csv', '.log', '.rtf', '.tex', '.bib']
    };
    let finalReport = `Gemini 分析报告\n==================\n\n用户指令: ${userPrompt}\n\n`;
    let fileCount = 0;
    const BATCH_SIZE_LIMIT_BYTES = 3.5 * 1024 * 1024; // 3.5MB limit, leaving headroom for API overhead
    let currentBatchParts = [];
    let currentBatchSizeBytes = 0;
    const sizeEstimationFactor = 1.5; // Factor to account for JSON overhead when estimating size

    for (const [filePath, entry] of Object.entries(entries)) {
      if (entry.isDirectory) {
          console.log("Skipping directory:", filePath);
          continue;
      }
      fileCount++;
      console.log("Processing file:", filePath);
      const fileData = new Uint8Array(await entry.arrayBuffer());
      if (fileData.length === 0) {
          console.log("Skipping empty file:", filePath);
          continue;
      }
      const extension = `.${filePath.split('.').pop()?.toLowerCase()}` || '';
      const contentHeader = `\n\n--- 文件: ${filePath} (${fileData.length} bytes) ---\n`;
      let partSegments = [];

      // Determine processing method based on file type
      if (fileTypes['图像 (Image)'].includes(extension)) {
          const mimeType = getMimeType(extension);
          if (!mimeType.startsWith('image/')) {
              console.warn("Skipping non-image file treated as image:", filePath);
              partSegments = [{ text: contentHeader + "[非标准图像文件，已跳过。]" }];
          } else {
              try {
                  const base64 = arrayBufferToBase64(fileData.buffer);
                  // Gemini API has size limits for image data, ensure base64 string is not too large
                  if (base64.length > 4 * 1024 * 1024) { // 4MB check for base64 string
                      console.warn("Image file too large for Gemini API, skipping:", filePath);
                      partSegments = [{ text: contentHeader + "[图像文件过大，无法分析。]" }];
                  } else {
                      partSegments = [{ text: contentHeader }, { inlineData: { mimeType, data: base64 } }];
                  }
              } catch (e) {
                  console.error("Error processing image:", filePath, e);
                  partSegments = [{ text: contentHeader + `[图像处理错误: ${e.message}]` }];
              }
          }
      } else if (fileTypes['核心文本/代码 (Core Text/Code)'].includes(extension) || fileTypes['其他文本 (Other Text)'].includes(extension)) {
           try {
              // Attempt to decode as UTF-8
              let textContent = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(fileData);
              // Sanitize content to remove potentially problematic control characters
              textContent = sanitizeText(textContent);
              // Check length to avoid passing massive files in one go, even if they are text
              if (textContent.length > 100000) { // Example: 100k chars max per file
                  console.warn("Text file too large, truncating for analysis:", filePath);
                  textContent = textContent.substring(0, 100000) + "\n... [内容已被截断以符合API限制] ...";
              }
              partSegments = [{ text: contentHeader + textContent }];
            } catch (e) {
              console.warn("File not UTF-8 decodable, skipping text analysis:", filePath, e.message);
              partSegments = [{ text: contentHeader + `[文件内容无法以UTF-8解码，已跳过。错误: ${e.message}]` }];
            }
      } else {
          // Unsupported file type
          partSegments = [{ text: contentHeader + "[不支持分析的文件类型，已跳过。]" }];
          console.log("Skipped unsupported file type:", filePath, extension);
      }

      // Estimate the size of this part for batching purposes
      // JSON.stringify is an approximation; actual API payload might be slightly different
      const partSizeEstimate = JSON.stringify(partSegments).length * sizeEstimationFactor;

      // Check if adding this part would exceed the batch size
      if (currentBatchSizeBytes + partSizeEstimate > BATCH_SIZE_LIMIT_BYTES && currentBatchParts.length > 0) {
        console.log("Batch size limit reached, processing current batch...");
        finalReport += await processBatch(currentBatchParts, userPrompt, env);
        currentBatchParts = [];
        currentBatchSizeBytes = 0;
      }
      
      // Add the current part to the batch
      currentBatchParts.push(...partSegments);
      currentBatchSizeBytes += partSizeEstimate;
      console.log(`Added file ${filePath} to batch. Batch size estimate: ${currentBatchSizeBytes.toFixed(2)} / ${BATCH_SIZE_LIMIT_BYTES} bytes`);
    }

    // Process the final batch if it contains any data
    if (currentBatchParts.length > 0) {
      console.log("Processing final batch...");
      finalReport += await processBatch(currentBatchParts, userPrompt, env);
    }
    
    finalReport += `\n\n==================\n分析完成。共处理 ${fileCount} 个有效文件。\nGenerated on: ${new Date().toISOString()}`;
    
    console.log("Analysis complete, saving report to OneDrive at:", resultKey);
    // Save the final report back to OneDrive
    const reportUploadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
    const reportResponse = await fetch(reportUploadUrl, {
        method: 'PUT',
        headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'text/plain; charset=utf-8' // Explicitly set charset
        },
        body: finalReport,
    });

    if (!reportResponse.ok) {
        const reportError = await reportResponse.text();
        console.error("Failed to save report to OneDrive:", reportResponse.status, reportResponse.statusText, reportError);
        throw new Error(`Failed to save report to OneDrive: ${reportResponse.status} ${reportResponse.statusText}. ${reportError}`);
    }

    console.log("Report successfully saved to OneDrive:", resultKey);

  } catch (e) {
    console.error("后台分析失败 for item", driveItemId, ":", e);
    // Attempt to save an error report to OneDrive, but handle token errors gracefully
    try {
        const accessToken = await getGraphApiAccessToken(env);
        if (accessToken) {
            const userId = env.MS_USER_ID || 'me';
            const errorReportUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${resultKey}:/content`;
            const errorReportContent = `分析失败: ${e.message}\n\n详细错误信息:\n${e.stack}\n\n请检查您的文件或联系支持。\nGenerated on: ${new Date().toISOString()}`;
            const errorResponse = await fetch(errorReportUrl, {
                method: 'PUT',
                headers: { 
                    'Authorization': `Bearer ${accessToken}`, 
                    'Content-Type': 'text/plain; charset=utf-8' // Ensure UTF-8 encoding
                },
                body: errorReportContent,
            });

            if (!errorResponse.ok) {
                console.error("Failed to save error report to OneDrive:", await errorResponse.text());
                // Do not throw here as the main analysis already failed
            } else {
                console.log("Error report saved to OneDrive:", resultKey);
            }
        } else {
            console.error("Could not save error report: No access token available.");
        }
    } catch (tokenError) {
        console.error("Could not save error report due to token error:", tokenError);
        // Do not throw here as the main analysis already failed
    }
  }
}

/**
 * Sends a batch of file parts to the Gemini API for analysis.
 */
async function processBatch(parts, userPrompt, env) {
    console.log("Sending batch to Gemini API, number of parts:", parts.length);
    const GOOGLE_API_KEY = env.GOOGLE_API_KEY;
    if (!GOOGLE_API_KEY) {
        console.error("GOOGLE_API_KEY environment variable is not set.");
        throw new Error("未配置 'GOOGLE_API_KEY' 环境变量。");
    }

    const model = 'gemini-2.5-flash'; // Or 'gemini-2.0-flash' or 'gemini-1.5-pro'
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    
    // Construct the payload for the Gemini API
    // The user prompt is the first part, followed by the file parts
    const contentParts = [{ text: userPrompt }, ...parts];
    const payload = { 
        contents: [{ role: "user", parts: contentParts }], // Explicitly set role as 'user'
        generationConfig: {
             // Optional: Add parameters like temperature, maxOutputTokens if needed
             // temperature: 0.5,
             // maxOutputTokens: 2048,
        }
    };

    try {
        const apiResponse = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        if (!apiResponse.ok) {
            const errorBody = await apiResponse.text();
            console.error("Gemini API Error:", apiResponse.status, apiResponse.statusText, errorBody);
            
            // Try to parse error details if possible
            let errorMessage = `API Error ${apiResponse.status}: ${apiResponse.statusText}`;
            try {
                const errorJson = JSON.parse(errorBody);
                if (errorJson.error && errorJson.error.message) {
                     errorMessage += `. Details: ${errorJson.error.message}`;
                }
            } catch (e) {
                // If parsing fails, use the raw text
                errorMessage += `. Raw response: ${errorBody.substring(0, 200)}...`; // Truncate long errors
            }
            
            return `\n[分析批次失败: ${errorMessage}]\n`;
        }

        const resultJson = await apiResponse.json();
        
        // Check for safety reasons or other issues in the response
        if (!resultJson.candidates || resultJson.candidates.length === 0) {
            console.warn("Gemini API returned no candidates:", resultJson);
            const blockReason = resultJson.promptFeedback?.blockReason || 'No candidates returned';
            return `\n[模型未返回有效内容或因安全原因被阻止: ${blockReason}]\n`;
        }
        
        const generatedText = resultJson.candidates[0]?.content?.parts?.[0]?.text;
        if (typeof generatedText !== 'string' || generatedText.trim() === '') {
            console.warn("Gemini API returned candidate but no valid text:", resultJson.candidates[0]);
            return "\n[模型返回了结果，但未包含有效文本内容]\n";
        }

        console.log("Gemini API batch processed successfully.");
        return generatedText;

    } catch (e) {
         console.error("Network or processing error calling Gemini API:", e);
         return `\n[调用 Gemini API 时发生网络或处理错误: ${e.message}]\n`;
    }
}

// --- Helper Functions ---
function getMimeType(extension) {
    const mimeTypes = { 
        '.png': 'image/png', 
        '.jpg': 'image/jpeg', 
        '.jpeg': 'image/jpeg', 
        '.gif': 'image/gif', 
        '.webp': 'image/webp',
        '.bmp': 'image/bmp',
        '.svg': 'image/svg+xml',
        '.tiff': 'image/tiff',
        '.ico': 'image/x-icon'
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
  // Remove null bytes, control characters (except common whitespace), and replace with a space or remove
  // Keeps \n, \r, \t
  return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, ' '); 
  // Alternatively, to remove them: .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
}


// --- HTML User Interface ---
const html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gemini 大型文件分析器 (OneDrive版)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f3f4f6; }
        .container { max-width: 800px; }
        #result-box, #status-box {
            white-space: pre-wrap; word-wrap: break-word; background-color: #1e293b;
            color: #e2e8f0; border-radius: 0.5rem; padding: 1.5rem; margin-top: 1.5rem;
            line-height: 1.75; font-family: 'Courier New', Courier, monospace;
            overflow-x: auto; /* Allow horizontal scrolling for long lines */
        }
        .loader {
            display: inline-block;
            border: 4px solid #f3f3f3; border-top: 4px solid #3b82f6; border-radius: 50%;
            width: 20px; height: 20px; animation: spin 1s linear infinite; margin-right: 10px; vertical-align: middle;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .file-input::file-selector-button {
            margin-right: 1rem;
            background: #2563eb;
            color: white;
            border: none;
            border-radius: 0.375rem;
            padding: 0.5rem 1rem;
            cursor: pointer;
        }
        .file-input::file-selector-button:hover {
            background: #1d4ed8;
        }
    </style>
</head>
<body class="antialiased text-gray-800">
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">
        <header class="text-center mb-8">
            <h1 class="text-3xl sm:text-4xl font-bold text-gray-900">Gemini 大型文件分析器</h1>
            <p class="mt-2 text-lg text-gray-600">由 OneDrive & Gemini 强力驱动</p>
        </header>
        <form id="upload-form" class="bg-white p-6 rounded-lg shadow-md border border-gray-200">
            <div class="mb-5">
                <label for="user_prompt" class="block mb-2 text-md font-medium text-gray-700">你的分析指令:</label>
                <textarea 
                    id="user_prompt" 
                    name="user_prompt" 
                    rows="4" 
                    class="w-full p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                    placeholder="例如：请总结每个代码文件的主要功能和潜在问题。"
                    required
                ></textarea>
            </div>
            <div class="mb-5">
                 <label for="zip-upload" class="block mb-2 text-md font-medium text-gray-700">选择一个 ZIP 压缩包上传 (最大 100MB):</label>
                 <!-- Added max file size hint and accept -->
                <input 
                    type="file" 
                    name="file" 
                    id="zip-upload" 
                    accept=".zip" 
                    class="file-input w-full text-sm text-gray-500
                          file:mr-4 file:py-2 file:px-4
                          file:rounded-md file:border-0
                          file:text-sm file:font-semibold
                          "
                    required
                />
            </div>
            <button 
                type="submit" 
                id="submit-btn" 
                class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-4 rounded-md transition duration-200 disabled:opacity-50 disabled:cursor-not-allowed"
            >
                开始分析
            </button>
        </form>
        <div id="status-container" class="mt-6"></div>
    </div>
    <script>
        const form = document.getElementById('upload-form');
        const submitBtn = document.getElementById('submit-btn');
        const zipUpload = document.getElementById('zip-upload');
        const statusContainer = document.getElementById('status-container');
        const userPrompt = document.getElementById('user_prompt');

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const file = zipUpload.files[0];
            if (!file) {
                updateStatus('<p class="text-red-500 font-semibold">错误：请选择一个 ZIP 文件。</p>', 'error');
                return;
            }
            // Optional: Client-side file size check (can be bypassed, server check is essential)
            if (file.size > 100 * 1024 * 1024) { // 100MB
                 updateStatus('<p class="text-red-500 font-semibold">错误：文件过大（超过100MB），请上传较小的文件。</p>', 'error');
                 return;
            }
            setLoadingState(true, '正在上传文件并启动分析... 这可能需要几分钟时间，请勿关闭页面。');
            try {
                const formData = new FormData();
                formData.append('file', file); // Use 'file' to match server-side get('file')
                formData.append('userPrompt', userPrompt.value);
                const analysisResponse = await fetch('/api/start-analysis', {
                    method: 'POST',
                    body: formData, // Send FormData directly
                });
                
                const responseText = await analysisResponse.text(); // Always read response text for error details
                
                if (!analysisResponse.ok) {
                    let errorMessage = '启动分析任务失败。';
                    try {
                        // Try to parse as JSON for structured error
                        const errorData = JSON.parse(responseText);
                        errorMessage = errorData.error || errorMessage;
                    } catch (e) {
                        // If not JSON, use the raw text
                        errorMessage = responseText || errorMessage;
                    }
                    throw new Error(errorMessage);
                }
                
                const analysisData = JSON.parse(responseText); // Parse successful JSON response
                updateStatus('文件上传成功！分析任务已在后台运行。正在等待最终报告...', 'loading');
                pollForResult(analysisData.resultKey);
            } catch (error) {
                console.error('流程错误:', error);
                updateStatus('发生错误：' + error.message, 'error');
                setLoadingState(false);
            }
        });

        function pollForResult(resultKey) {
            const pollInterval = 15000; // 15 seconds
            const maxAttempts = 120; // 120 * 15s = 30 minutes
            let attempts = 0;
            const intervalId = setInterval(async () => {
                if (attempts++ > maxAttempts) {
                    clearInterval(intervalId);
                    updateStatus('分析超时（30分钟）。请检查文件或联系支持。', 'error');
                    setLoadingState(false);
                    return;
                }
                console.log('Polling for result, attempt', attempts, 'for key:', resultKey);
                try {
                    const resultResponse = await fetch('/' + resultKey);
                    if (resultResponse.status === 200) {
                        clearInterval(intervalId);
                        const reportText = await resultResponse.text();
                        updateStatus('分析完成！', 'success');
                        displayResult(reportText, resultKey);
                        setLoadingState(false);
                    } else if (resultResponse.status === 404) {
                        // Still processing, continue polling
                        console.log('Result not ready yet (404), continuing to poll.');
                    } else {
                        // Other error occurred
                        clearInterval(intervalId);
                        const errorText = await resultResponse.text();
                        throw new Error('获取结果时发生服务器错误: ' + resultResponse.status + ' ' + errorText);
                    }
                } catch (error) {
                    clearInterval(intervalId);
                    console.error('轮询错误:', error);
                    updateStatus('获取结果失败：' + error.message, 'error');
                    setLoadingState(false);
                }
            }, pollInterval);
        }

        function setLoadingState(isLoading, message = '') {
            submitBtn.disabled = isLoading;
            submitBtn.textContent = isLoading ? '处理中...' : '开始分析';
            if (isLoading) {
                updateStatus(message, 'loading');
            }
        }

        function updateStatus(message, type) {
            let content = '';
            if (type === 'loading') {
                content = '<div id="status-box"><div class="loader"></div><span>' + message + '</span></div>';
            } else if (type === 'error') {
                content = '<div id="status-box"><p class="text-red-500 font-semibold">' + message + '</p></div>';
            } else { // success
                 content = '<div id="status-box"><p class="text-green-500 font-semibold">' + message + '</p></div>';
            }
            statusContainer.innerHTML = content;
        }

        function displayResult(reportText, resultKey) {
            // Sanitize reportText to prevent XSS, though server should ideally handle this
            const sanitizedText = reportText
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;");

            const resultHtml = 
                '<div id="result-box">' +
                    '<a href="/' + resultKey + '" target="_blank" class="float-right bg-gray-600 hover:bg-gray-500 text-white font-bold py-1 px-3 rounded-md text-sm transition duration-200 no-underline">在新标签页打开报告</a>' + // Changed from download to view in new tab
                    '<pre>' + sanitizedText + '</pre>' + // Use <pre> to preserve formatting
                '</div>';
            statusContainer.innerHTML += resultHtml;
        }
    </script>
</body>
</html>
`;
}



