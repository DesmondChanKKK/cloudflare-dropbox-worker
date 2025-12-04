import * as XLSX from 'xlsx';

const VERSION = '1.0.1';

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

    
        // Validate Client ID
        if (clientId !== env.DROPBOX_APP_KEY) {
            console.error(`[Debug] Unauthorized: Invalid clientid.`);
            return createResponse({ error: 'Unauthorized: Invalid clientid.' }, 401);
        }

        if (!filename) {
            console.error('[Debug] Error: No filename provided');
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
            console.log('[Debug] Attempting to use Refresh Token flow...');
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
                    console.error(`[Debug] Refresh Token Error: ${tokenResponse.status} ${errorText}`);
                    return createResponse({ error: `Dropbox Auth Error: ${tokenResponse.status} ${errorText}` }, 500);
                }

                const tokenData = await tokenResponse.json();
                accessToken = tokenData.access_token;
                console.log('[Debug] Successfully refreshed Access Token');
            } catch (error) {
                console.error(`[Debug] Auth Exception: ${error.message}`);
                return createResponse({ error: `Auth Flow Error: ${error.message}` }, 500);
            }
        } else {
            console.log('[Debug] Using static Access Token');
        }

        // Priority 2: Check if we have ANY token
        if (!accessToken) {
            console.error('[Debug] Error: No access token available');
            return createResponse({ error: 'Configuration Error: No valid Dropbox Token found. Please set DROPBOX_ACCESS_TOKEN or (DROPBOX_REFRESH_TOKEN + APP_KEY + APP_SECRET).' }, 500);
        }

        try {
            // 1. Fetch file from Dropbox
            const dropboxResponse = await fetch('https://content.dropboxapi.com/2/files/download', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Dropbox-API-Arg': JSON.stringify({
                        path: dropboxPath
                    })
                }
            });

            if (!dropboxResponse.ok) {
                const errorText = await dropboxResponse.text();
                console.error(`[Debug] Dropbox Download Error: ${dropboxResponse.status} ${errorText}`);
                return createResponse({ error: `Dropbox Download Error: ${dropboxResponse.status} ${errorText}. Requested Path: ${dropboxPath}` }, dropboxResponse.status);
            }

            // 2. Parse Excel file
            console.log('[Debug] File downloaded, parsing Excel...');
            const arrayBuffer = await dropboxResponse.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

            // Assuming we want the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to array of arrays to handle layout better
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Handle 'raw' request type - return full data immediately
            if (requestType === 'raw') {
                console.log('[Debug] Request type is "raw", returning full dataset.');
                return createResponse({ data: rows });
            }

            // Initialize result object
            const result = {
                currency: 'EUR' // Default assumption
            };

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
                    console.error(`[Debug] Failed to parse EXTRACTION_CONFIG: ${e.message}`);
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
                            console.warn(`[Debug] JSON parse failed for body: ${e.message}`);
                        }
                    }

                    if (bodyConfig && Array.isArray(bodyConfig)) {
                        extractionRules = bodyConfig;
                    }
                } catch (e) {
                    console.error(`[Debug] Failed to process request body: ${e.message}`);
                }

                // 2. Fallback to Query Parameter
                if (extractionRules.length === 0) {
                    if (customConfigStr) {
                        try {
                            extractionRules = JSON.parse(customConfigStr);
                        } catch (e) {
                            console.error(`[Debug] Failed to parse custom 'config' parameter: ${e.message}`);
                            return createResponse({ error: 'Invalid JSON in "config" parameter.' }, 400);
                        }
                    } else {
                        console.error('[Debug] Type is "custom" but no configuration provided.');
                        return createResponse({ error: 'Type is "custom" but missing configuration in Body or "config" query parameter.' }, 400);
                    }
                }
            } else {
                // Look up in loaded configs
                extractionRules = allConfigs[requestType] || allConfigs['default'];

                if (!extractionRules) {
                    console.warn(`[Debug] No configuration found for type "${requestType}", and no default available.`);
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

            // Iterate through rows
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                // Check both Column A (index 0) and Column B (index 1) for keywords
                const cellA = row[0] ? row[0].toString().toLowerCase() : '';
                const cellB = row[1] ? row[1].toString().toLowerCase() : '';
                const textToCheck = cellA + ' ' + cellB; // Combine them to be safe

                // Check against each rule
                for (const rule of extractionRules) {
                    // Check if ALL keywords in the rule are present in the text
                    const match = rule.keywords.every(keyword => textToCheck.includes(keyword.toLowerCase()));

                    if (match) {
                        const val = row[rule.colIndex];
                        const parsedVal = parseCurrency(val);
                        result[rule.key] = parsedVal;
                    }
                }
            }

            // 3. Return Structured JSON
            return createResponse(result);

        } catch (error) {
            console.error(`[Debug] Worker Exception: ${error.message}`);
            return createResponse({ error: `Worker Error: ${error.message}` }, 500);
        }
    }
};
