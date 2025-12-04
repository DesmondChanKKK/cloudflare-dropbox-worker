const fetch = require('node-fetch'); // Ensure you have node-fetch or use Node 18+

// Configuration
const WORKER_URL = 'http://127.0.0.1:8787'; // Default local wrangler dev URL
// const WORKER_URL = 'https://your-worker.workers.dev'; // Uncomment and set for production

async function testWorker(filename, type, folder, customConfig) {
    const url = new URL(WORKER_URL);
    url.searchParams.set('filename', filename);
    if (type) {
        url.searchParams.set('type', type);
    }
    if (folder) {
        url.searchParams.set('folder', folder);
    }
    if (customConfig) {
        url.searchParams.set('config', JSON.stringify(customConfig));
    }

    console.log(`\n[Request] ${url.toString()}`);

    try {
        const response = await fetch(url);
        const text = await response.text();

        console.log(`[Status] ${response.status} ${response.statusText}`);

        try {
            const json = JSON.parse(text);
            console.log('[Response Body (JSON)]:', JSON.stringify(json, null, 2));
        } catch {
            console.log('[Response Body (Text)]:', text);
        }
    } catch (error) {
        console.error('[Error]:', error.message);
    }
}

// Main execution
(async () => {
    // Example 1: Default config, Root folder
    await testWorker('test_file.xlsx');

    // Example 2: Specific config, Root folder
    await testWorker('test_file.xlsx', 'example_type');

    // Example 3: Default config, Specific folder
    await testWorker('test_file.xlsx', null, '2023/reports');

    // Example 4: Custom Configuration (Dynamic)
    const myCustomRules = [
        { key: "dynamic_total", keywords: ["dynamic", "total"], colIndex: 4 }
    ];
    await testWorker('test_file.xlsx', 'custom', null, myCustomRules);
})();
