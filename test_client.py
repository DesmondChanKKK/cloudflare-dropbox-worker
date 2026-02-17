import requests
import json
import argparse
import sys

# 默认配置 (你可以修改这里，或者通过命令行参数传入)
DEFAULT_WORKER_URL = "https://cloudflare-dropbox-worker.kong-smartway.workers.dev" # 生产环境地址
# DEFAULT_WORKER_URL = "http://127.0.0.1:8787" # 本地调试默认地址

# 必须与 wrangler.toml 中的 DROPBOX_APP_KEY 一致
DEFAULT_CLIENT_ID = "7t18hxy1q3fj3tv" 

def test_worker(url, client_id, filename, folder=None, request_type="default", config=None):
    """
    发送请求给 Worker 并打印结果
    """
    params = {
        "filename": filename,
        "clientid": client_id,
        "type": request_type
    }
    
    if folder:
        params["folder"] = folder
        
    if config:
        params["config"] = json.dumps(config)
        
    print(f"\n--- Request Details ---")
    print(f"URL: {url}")
    print(f"Params: {json.dumps(params, ensure_ascii=False)}")
    
    try:
        response = requests.get(url, params=params)
        
        print(f"\n--- Response ({response.status_code}) ---")
        
        if response.status_code == 200:
            try:
                data = response.json()
                print(json.dumps(data, indent=2, ensure_ascii=False))
            except json.JSONDecodeError:
                print("Response is not JSON:")
                print(response.text)
        else:
            print(f"Error: {response.text}")
            
    except requests.exceptions.RequestException as e:
        print(f"Connection Error: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test Cloudflare Dropbox Worker")
    parser.add_argument("--url", default=DEFAULT_WORKER_URL, help="Worker URL")
    parser.add_argument("--key", default=DEFAULT_CLIENT_ID, help="Client ID (App Key)")
    parser.add_argument("--file", required=True, help="Filename in Dropbox")
    parser.add_argument("--folder", help="Folder in Dropbox")
    parser.add_argument("--type", default="default", help="Extraction type (default, custom, raw, etc.)")
    parser.add_argument("--config", help="JSON string for custom config (only used if type=custom)")
    
    # Check if any args are passed, if not, allow interactive mode or demo
    if len(sys.argv) == 1:
        print("No arguments provided. Running demo mode...")
        print(f"Target: {DEFAULT_WORKER_URL}")
        
        # Demo 1: Default
        print("\n=== Demo 1: Default Extraction ===")
        test_worker(DEFAULT_WORKER_URL, DEFAULT_CLIENT_ID, "test.xlsx")
        
        # Demo 2: Custom Config
        print("\n=== Demo 2: Custom Extraction ===")
        custom_rules = [
            {"key": "demo_total", "keywords": ["Total", "总计"], "colIndex": 3}
        ]
        test_worker(DEFAULT_WORKER_URL, DEFAULT_CLIENT_ID, "test.xlsx", request_type="custom", config=custom_rules)
        
    else:
        args = parser.parse_args()
        
        config = None
        if args.config:
            try:
                config = json.loads(args.config)
            except json.JSONDecodeError:
                print("Error: --config must be valid JSON string")
                sys.exit(1)
                
        test_worker(args.url, args.key, args.file, args.folder, args.type, config)
