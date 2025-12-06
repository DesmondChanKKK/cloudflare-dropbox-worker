import * as XLSX from 'xlsx';

const VERSION = '1.1.1';
const DEFAULT_IS_DEBUG = true; // Default to true as requested, can be overridden by env

function createResponse(data, status = 200) {
    const body = {
        version: VERSION,
        ...(typeof data === 'string' ? { message: data } : data)
    };
    return new Response(JSON.stringify(body), {
        status: status,
        headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
        }
    });
}

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);
        const filename = url.searchParams.get('filename');
        const folder = url.searchParams.get('folder') || ''; // Default to empty if not provided
        const requestType = url.searchParams.get('type') || 'default';
        const customConfigStr = url.searchParams.get('config');
        const clientId = url.searchParams.get('clientid');

        // Determine Debug Mode
        // Check environment variable first, then fallback to default
        const isDebug = env.IS_DEBUG !== undefined ? (env.IS_DEBUG === 'true') : DEFAULT_IS_DEBUG;

        const logger = {
            debug: (...args) => {
                if (isDebug) console.log(...args);
            },
            error: (...args) => {
                // Always log errors, or maybe only in debug? 
                // User said "if true, print test DEBUG logs". 
                // Critical errors should usually be printed. 
                // But for "Debug" prefixed errors that are just info for dev, we can check isDebug.
                const msg = args.join(' ');
                if (msg.startsWith('[Debug]') && !isDebug) return;
                console.error(...args);
            },
            warn: (...args) => {
                if (isDebug) console.warn(...args);
            }
        };

        // Validate Client ID
        if (!env.DROPBOX_APP_KEY) {
            logger.error('[Debug] Configuration Error: DROPBOX_APP_KEY is missing.');
            return createResponse({ error: 'Configuration Error: DROPBOX_APP_KEY is not set on the server. Did you forget to set it as a secret?' }, 500);
        }

        if (clientId !== env.DROPBOX_APP_KEY) {
            logger.error(`[Debug] Unauthorized: Invalid clientid.`);
            return createResponse({ error: 'Unauthorized: Invalid clientid.' }, 401);
        }

        if (!filename) {
            logger.error('[Debug] Error: No filename provided');
            return createResponse({ error: 'Please provide a "filename" query parameter.' }, 400);
        }

        // Construct Dropbox Path
        // Ensure folder starts with / if it exists, and doesn't end with /
        let dropboxPath = '';
        if (folder) {
            const cleanFolder = folder.startsWith('/') ? folder : `/${folder}`;
            const finalFolder = cleanFolder.endsWith('/') ? cleanFolder.slice(0, -1) : cleanFolder;
            dropboxPath = `${finalFolder}/${filename}`;
        } else {
            dropboxPath = `/${filename}`;
        }

        // Sanitize path
        dropboxPath = dropboxPath.replace(/\/+/g, '/');

        let accessToken = env.DROPBOX_ACCESS_TOKEN;

        // Priority 1: Try Refresh Token Flow
        if (env.DROPBOX_REFRESH_TOKEN && env.DROPBOX_APP_KEY && env.DROPBOX_APP_SECRET) {
            try {
                const tokenResponse = await fetch('https://api.dropbox.com/oauth2/token', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: new URLSearchParams({
                        grant_type: 'refresh_token',
                        refresh_token: env.DROPBOX_REFRESH_TOKEN,
                        client_id: env.DROPBOX_APP_KEY,
                        client_secret: env.DROPBOX_APP_SECRET,
                    }),
                });

                if (!tokenResponse.ok) {
                    const errorText = await tokenResponse.text();
                    logger.error(`[Debug] Refresh Token Error: ${tokenResponse.status} ${errorText}`);
                    return createResponse({ error: `Dropbox Auth Error: ${tokenResponse.status} ${errorText}` }, 500);
                }

                const tokenData = await tokenResponse.json();
                accessToken = tokenData.access_token;
            } catch (error) {
                logger.error(`[Debug] Auth Exception: ${error.message}`);
                return createResponse({ error: `Auth Flow Error: ${error.message}` }, 500);
            }
        } else {
            logger.debug('[Debug] Using static Access Token');
        }

        // Priority 2: Check if we have ANY token
        if (!accessToken) {
            logger.error('[Debug] Error: No access token available');
            return createResponse({ error: 'Configuration Error: No valid Dropbox Token found. Please set DROPBOX_ACCESS_TOKEN or (DROPBOX_REFRESH_TOKEN + APP_KEY + APP_SECRET).' }, 500);
        }

        try {
            // Helper function to download file
            const downloadFile = async (path) => {
                const response = await fetch('https://content.dropboxapi.com/2/files/download', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Dropbox-API-Arg': JSON.stringify({
                            path: path
                        })
                    }
                });
                return response;
            };

            // Helper function to search file recursively
            const searchFileRecursively = async (searchFolder, searchFilename) => {
                logger.debug(`[Debug] Searching recursively for "${searchFilename}" in "${searchFolder}"...`);
                let hasMore = true;
                let cursor = null;
                const lowerSearchFilename = searchFilename.toLowerCase();

                try {
                    while (hasMore) {
                        let url = 'https://api.dropboxapi.com/2/files/list_folder';
                        let body = {
                            path: searchFolder === '/' ? '' : searchFolder,
                            recursive: true,
                            include_media_info: false,
                            include_deleted: false,
                            include_has_explicit_shared_members: false,
                            include_mounted_folders: true,
                            limit: 2000
                        };

                        if (cursor) {
                            url = 'https://api.dropboxapi.com/2/files/list_folder/continue';
                            body = { cursor: cursor };
                        }

                        const response = await fetch(url, {
                            method: 'POST',
                            headers: {
                                'Authorization': `Bearer ${accessToken}`,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify(body)
                        });

                        if (!response.ok) {
                            const errorText = await response.text();
                            logger.error(`[Debug] List Folder Error: ${response.status} ${errorText}`);
                            return null;
                        }

                        const data = await response.json();

                        // Log all folders found in this batch
                        if (isDebug) {
                            data.entries.forEach(entry => {
                                if (entry['.tag'] === 'folder') {
                                    logger.debug(`[Debug] Scanned folder: ${entry.path_display}`);
                                }
                            });
                        }

                        // Find match (Case-Insensitive)
                        const match = data.entries.find(entry => 
                            entry['.tag'] === 'file' && entry.name.toLowerCase() === lowerSearchFilename
                        );

                        if (match) {
                            return match.path_lower;
                        }

                        hasMore = data.has_more;
                        cursor = data.cursor;
                    }

                    return null;
                } catch (e) {
                    logger.error(`[Debug] Recursive search failed: ${e.message}`);
                    return null;
                }
            };

            // 1. Attempt to download file from initial path
            let dropboxResponse = await downloadFile(dropboxPath);

            // 2. If not found (409), try recursive search
            if (dropboxResponse.status === 409) {
                logger.debug(`[Debug] File not found at ${dropboxPath}. Attempting recursive search...`);
                
                // Determine the base folder to search in. 
                // If 'folder' param was provided, search inside that. 
                // If not, search root (which might be expensive/slow but requested).
                const baseSearchFolder = folder ? (folder.startsWith('/') ? folder : `/${folder}`) : '';
                
                const foundPath = await searchFileRecursively(baseSearchFolder, filename);

                if (foundPath) {
                    logger.debug(`[Debug] File found at ${foundPath}. Retrying download...`);
                    dropboxResponse = await downloadFile(foundPath);
                } else {
                    logger.debug(`[Debug] File "${filename}" not found recursively in "${baseSearchFolder}".`);
                }
            }

            if (!dropboxResponse.ok) {
                const errorText = await dropboxResponse.text();
                logger.error(`[Debug] Dropbox Download Error: ${dropboxResponse.status} ${errorText}`);
                return createResponse({ error: `Dropbox Download Error: ${dropboxResponse.status} ${errorText}. Requested Path: ${dropboxPath}` }, dropboxResponse.status);
            }

            // 3. Parse Excel file
            logger.debug('[Debug] File downloaded, parsing Excel...');
            const arrayBuffer = await dropboxResponse.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

            // Assuming we want the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to array of arrays to handle layout better
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Handle 'raw' request type - return full data immediately
            if (requestType === 'raw') {
                logger.debug('[Debug] Request type is "raw", returning full dataset.');
                return createResponse({ data: rows });
            }



            // Default Configuration
            let allConfigs = {
                "default": [
                    { key: "hardware_total", keywords: ["hardware", "小计"], colIndex: 3 },
                    { key: "service_total", keywords: ["service", "小计"], colIndex: 3 },
                    { key: "grand_total", keywords: ["总合计"], colIndex: 3 },
                    { key: "grand_total", keywords: ["iva inclusa"], colIndex: 3 }
                ]
            };

            // Load Config from Environment Variable if present
            if (env.EXTRACTION_CONFIG) {
                try {
                    const parsedConfig = JSON.parse(env.EXTRACTION_CONFIG);
                    if (Array.isArray(parsedConfig)) {
                        // Support legacy array format by assigning it to 'default'
                        allConfigs["default"] = parsedConfig;
                    } else {
                        // Assume it's the new object format
                        allConfigs = { ...allConfigs, ...parsedConfig };
                    }
                } catch (e) {
                    logger.error(`[Debug] Failed to parse EXTRACTION_CONFIG: ${e.message}`);
                }
            }

            // Determine Extraction Rules
            let extractionRules = [];

            if (requestType === 'custom') {
                // 1. Try to read from Body
                try {
                    const rawBodyText = await request.clone().text();
                    let bodyConfig = null;
                    
                    if (rawBodyText) {
                        try {
                            bodyConfig = JSON.parse(rawBodyText);
                            // Handle double-encoded JSON string
                            if (typeof bodyConfig === 'string') {
                                try {
                                    bodyConfig = JSON.parse(bodyConfig);
                                } catch (e) {
                                    // Ignore
                                }
                            }
                        } catch (e) {
                            logger.warn(`[Debug] JSON parse failed for body: ${e.message}`);
                        }
                    }

                    if (bodyConfig && Array.isArray(bodyConfig)) {
                        extractionRules = bodyConfig;
                    }
                } catch (e) {
                    logger.error(`[Debug] Failed to process request body: ${e.message}`);
                }

                // 2. Fallback to Query Parameter
                if (extractionRules.length === 0) {
                    if (customConfigStr) {
                        try {
                            extractionRules = JSON.parse(customConfigStr);
                        } catch (e) {
                            logger.error(`[Debug] Failed to parse custom 'config' parameter: ${e.message}`);
                            return createResponse({ error: 'Invalid JSON in "config" parameter.' }, 400);
                        }
                    } else {
                        logger.error('[Debug] Type is "custom" but no configuration provided.');
                        return createResponse({ error: 'Type is "custom" but missing configuration in Body or "config" query parameter.' }, 400);
                    }
                }
            } else {
                // Look up in loaded configs
                extractionRules = allConfigs[requestType] || allConfigs['default'];

                if (!extractionRules) {
                    logger.warn(`[Debug] No configuration found for type "${requestType}", and no default available.`);
                    extractionRules = []; // Prevent crash
                }
            }

            // Helper function to clean and parse currency string
            const parseCurrency = (value) => {
                if (typeof value === 'number') return value;
                if (!value) return 0;
                const cleanStr = value.toString().replace(/[€$£\s]/g, '').replace(',', '.');
                return parseFloat(cleanStr) || 0;
            };

            // Helper function to extract data based on rules
            const extractData = (rows, rules) => {
                const extracted = {};
                // Initialize keys with 0
                rules.forEach(rule => extracted[rule.key] = 0);

                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    const cellA = row[0] ? row[0].toString().toLowerCase() : '';
                    const cellB = row[1] ? row[1].toString().toLowerCase() : '';
                    const textToCheck = cellA + ' ' + cellB;

                    for (const rule of rules) {
                        const match = rule.keywords.every(keyword => textToCheck.includes(keyword.toLowerCase()));
                        if (match) {
                            const val = row[rule.colIndex];
                            const parsedVal = parseCurrency(val);
                            // Only update if we found a non-zero value or if it's the first time?
                            // Logic implies last match wins, or we just take the first one?
                            // Original logic: result[rule.key] = parsedVal; (Last match wins if multiple rows match)
                            extracted[rule.key] = parsedVal;
                        }
                    }
                }
                return extracted;
            };

            // Perform Extraction
            let extractedData = extractData(rows, extractionRules);

            // Optimization for 'custom' type: Retry with colIndex + 1 if all results are 0
            if (requestType === 'custom') {
                const allZeros = Object.values(extractedData).every(val => val === 0);
                if (allZeros) {
                    logger.debug('[Debug] Custom extraction yielded all zeros. Retrying with colIndex + 1...');
                    const adjustedRules = extractionRules.map(rule => ({
                        ...rule,
                        colIndex: rule.colIndex + 1
                    }));
                    const retryData = extractData(rows, adjustedRules);
                    
                    // Check if retry yielded any non-zero results
                    const retryHasValues = Object.values(retryData).some(val => val !== 0);
                    if (retryHasValues) {
                        logger.debug('[Debug] Retry successful. Using adjusted column indices.');
                        extractedData = retryData;
                    } else {
                        logger.debug('[Debug] Retry also yielded all zeros. Reverting to original.');
                    }
                }
            }

            // Merge with default result structure
            const result = {
                currency: 'EUR', // Default assumption
                ...extractedData
            };

            // 3. Return Structured JSON
            return createResponse(result);

        } catch (error) {
            console.error(`[Debug] Worker Exception: ${error.message}`);
            return createResponse({ error: `Worker Error: ${error.message}` }, 500);
        }
    }
};
