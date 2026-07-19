import requests
import base64
import time
import json
import os

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
WORKFLOW_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"
IMAGE_PATH = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"
OUTPUT_DIR = "/home/newuser/TestAll/ocr_results"

os.makedirs(OUTPUT_DIR, exist_ok=True)

def upload_image():
    url = f"https://app.nanonets.com/api/v4/Workflow/{WORKFLOW_ID}/async/"
    headers = {
        "Authorization": f"Basic {base64.b64encode(f'{API_KEY}:'.encode()).decode()}"
    }
    with open(IMAGE_PATH, 'rb') as f:
        files = {"file": f}
        data = {"async": "true"}
        response = requests.post(url, headers=headers, files=files, data=data)
    return response.json()

def get_result(request_id):
    url = f"https://app.nanonets.com/api/v4/Workflow/{WORKFLOW_ID}/result/{request_id}/"
    headers = {
        "Authorization": f"Basic {base64.b64encode(f'{API_KEY}:'.encode()).decode()}"
    }
    response = requests.get(url, headers=headers)
    return response.json()

def main():
    print("Uploading image...")
    upload_response = upload_image()
    print(f"Upload response: {json.dumps(upload_response, indent=2)}")
    
    if "result_url" not in upload_response:
        print("Error: No result_url in response")
        return
    
    # Extract request ID from result_url? The response may contain request_id or result_url
    # According to Nanonets v4 async endpoint, the response should have a "result_url"
    # We can poll that URL directly.
    result_url = upload_response["result_url"]
    print(f"Result URL: {result_url}")
    
    # Poll for result
    while True:
        print("Polling for result...")
        result_response = requests.get(result_url, headers={
            "Authorization": f"Basic {base64.b64encode(f'{API_KEY}:'.encode()).decode()}"
        })
        result_data = result_response.json()
        print(f"Poll response: {json.dumps(result_data, indent=2)}")
        
        if result_data.get("status") == "done":
            print("OCR completed!")
            break
        elif result_data.get("status") == "failed":
            print("OCR failed!")
            return
        else:
            print("Still processing...")
            time.sleep(2)
    
    # Save the result JSON
    result_json = result_data.get("result", {})
    output_json = os.path.join(OUTPUT_DIR, f"result_{os.path.basename(IMAGE_PATH)}.json")
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(result_json, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON result to: {output_json}")
    
    # Optionally, we can also try to get markdown and docx if the API supports it?
    # But for now, we just save the JSON.

if __name__ == "__main__":
    main()
