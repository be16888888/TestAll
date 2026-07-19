import requests
import base64
import time
import json
import os
from pathlib import Path

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
WORKFLOW_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"
IMAGE_PATH = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"
OUTPUT_DIR = "/home/newuser/TestAll/ocr_results"

os.makedirs(OUTPUT_DIR, exist_ok=True)

AUTH = base64.b64encode(f"{API_KEY}:".encode()).decode()
HEADERS = {"Authorization": f"Basic {AUTH}"}

def upload_document():
    url = f"https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}/documents/"
    with open(IMAGE_PATH, 'rb') as f:
        files = {'file': (Path(IMAGE_PATH).name, f, 'image/webp')}
        data = {'async': 'false'}
        resp = requests.post(url, headers=HEADERS, files=files, data=data, timeout=120)
    resp.raise_for_status()
    return resp.json()

def get_document(document_id):
    url = f"https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}/documents/{document_id}"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.json()

def main():
    try:
        print("Uploading image...")
        upload_response = upload_document()
        print(f"Upload response: {json.dumps(upload_response, indent=2)}")
        
        document_id = upload_response.get('document_id')
        if not document_id:
            print("Error: No document_id in response")
            return
        
        # Save the upload response as initial result
        json_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_upload.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(upload_response, f, indent=2, ensure_ascii=False)
        print(f"Upload response saved to {json_path}")
        
        # Try to get the document after a short wait, but expect it might fail
        print(f"Waiting for processing (document_id: {document_id})...")
        max_wait = 30
        start = time.time()
        while time.time() - start < max_wait:
            try:
                doc = get_document(document_id)
                # Check if we have data
                has_data = False
                for page in doc.get('pages', []):
                    page_data = page.get('data', {})
                    if page_data.get('fields') or page_data.get('tables'):
                        has_data = True
                        break
                if has_data:
                    print("Processing complete with data!")
                    break
                else:
                    print("Still processing (no data yet)...")
            except Exception as e:
                print(f"Error fetching document: {e}")
            time.sleep(3)
        else:
            print("Timeout or persistent errors, using upload response.")
            doc = upload_response  # fallback
        
        # Save the final document response
        final_json_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_final.json")
        with open(final_json_path, 'w', encoding='utf-8') as f:
            json.dump(doc, f, indent=2, ensure_ascii=False)
        print(f"Final response saved to {final_json_path}")
        
        # Extract and save as markdown if there's data
        fields = {}
        items = []
        for page in doc.get('pages', []):
            page_data = page.get('data', {})
            for field in page_data.get('fields', []):
                label = field.get('label', '').strip()
                text = field.get('text', '').strip()
                if label and text:
                    fields[label] = text
            for table in page_data.get('tables', []):
                for row in table.get('rows', []):
                    cells = row.get('cells', [])
                    item = {}
                    for cell in cells:
                        label = cell.get('label', '').strip()
                        text = cell.get('text', '').strip()
                        if label and text:
                            item[label] = text
                    if item:
                        items.append(item)
        
        md_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_result.md")
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(f"# OCR Results - {Path(IMAGE_PATH).name}\n\n")
            f.write(f"**Workflow ID**: {WORKFLOW_ID}\n")
            f.write(f"**Processing Time**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write("## Extracted Fields\n")
            if fields:
                for k, v in fields.items():
                    f.write(f"- **{k}**: {v}\n")
            else:
                f.write("_No fields detected_\n")
            f.write("\n## Extracted Tables\n")
            if items:
                headers = list(items[0].keys())
                f.write("| " + " | ".join(headers) + " |\n")
                f.write("| " + " | ".join(["---"] * len(headers)) + " |\n")
                for item in items:
                    f.write("| " + " | ".join(item.get(h, "") for h in headers) + " |\n")
            else:
                f.write("_No table data detected_\n")
        print(f"Markdown report saved to {md_path}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
