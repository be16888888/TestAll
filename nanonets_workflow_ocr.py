import requests
import base64
import time
import json
import os
from pathlib import Path

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
BASE_URL = "https://app.nanonets.com/api/v4"
AUTH = base64.b64encode(f"{API_KEY}:".encode()).decode()
HEADERS = {"Authorization": f"Basic {AUTH}"}

def create_workflow():
    url = f"{BASE_URL}/workflows/"
    data = {
        "description": "Handwritten inventory table",
        "workflow_type": "tables"  # or "" for instant learning
    }
    resp = requests.post(url, headers=HEADERS, json=data, timeout=30)
    resp.raise_for_status()
    return resp.json()

def upload_document(workflow_id, image_path):
    url = f"{BASE_URL}/workflows/{workflow_id}/documents/upload"
    with open(image_path, 'rb') as f:
        files = {'file': (Path(image_path).name, f, 'image/webp')}
        data = {'async': 'false'}
        resp = requests.post(url, headers=HEADERS, files=files, data=data, timeout=120)
    resp.raise_for_status()
    return resp.json()

def get_document(workflow_id, document_id):
    url = f"{BASE_URL}/workflows/{workflow_id}/documents/{document_id}"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.json()

def extract_data(doc_json):
    fields = {}
    items = []
    for page in doc_json.get('pages', []):
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
    return fields, items

def save_markdown(fields, items, image_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    md_path = os.path.join(output_dir, f"{Path(image_path).stem}_result.md")
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# OCR Results - {Path(image_path).name}\n\n")
        f.write(f"**Workflow ID**: {workflow_id}\n")
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
    return md_path

def main():
    image_path = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"
    output_dir = "/home/newuser/TestAll/ocr_results"
    
    try:
        print("Creating workflow...")
        workflow = create_workflow()
        workflow_id = workflow.get('id')
        print(f"Created workflow: {workflow_id}")
        
        print(f"Uploading {image_path}...")
        upload_resp = upload_document(workflow_id, image_path)
        print(f"Upload response: {json.dumps(upload_resp, indent=2)}")
        
        document_id = upload_resp.get('document_id')
        if not document_id:
            print("Error: No document_id in response")
            return
        
        print(f"Waiting for processing (document_id: {document_id})...")
        while True:
            doc = get_document(workflow_id, document_id)
            # Check if processing is complete by looking for data
            ready = False
            for page in doc.get('pages', []):
                page_data = page.get('data', {})
                if page_data.get('fields') or page_data.get('tables'):
                    ready = True
                    break
            if ready:
                print("Processing complete!")
                break
            else:
                print("Still processing...")
                time.sleep(3)
        
        # Save raw JSON
        json_path = os.path.join(output_dir, f"{Path(image_path).stem}_raw.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(doc, f, indent=2, ensure_ascii=False)
        print(f"Raw JSON saved to {json_path}")
        
        fields, items = extract_data(doc)
        print(f"Extracted {len(fields)} fields and {len(items)} table rows")
        
        md_path = save_markdown(fields, items, image_path, output_dir)
        print(f"Markdown report saved to {md_path}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
