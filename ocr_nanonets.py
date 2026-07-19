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

def upload_document(workflow_id, image_path):
    url = f'https://app.nanonets.com/api/v4/workflows/{workflow_id}/documents/upload'
    auth_header = f"Basic {base64.b64encode(f'{API_KEY}:'.encode()).decode()}"
    
    with open(image_path, 'rb') as f:
        files = {'file': (Path(image_path).name, f, 'image/webp')}
        data = {'async': 'false'}  # synchronous processing
        
        response = requests.post(
            url,
            headers={'Authorization': auth_header},
            files=files,
            data=data,
            timeout=120
        )
    response.raise_for_status()
    return response.json()

def wait_for_completion(workflow_id, document_id, max_wait=120, interval=3):
    url = f'https://app.nanonets.com/api/v4/workflows/{workflow_id}/documents/{document_id}'
    auth_header = f"Basic {base64.b64encode(f'{API_KEY}:'.encode()).decode()}"
    start = time.time()
    
    while time.time() - start < max_wait:
        response = requests.get(url, headers={'Authorization': auth_header}, timeout=30)
        if response.status_code == 200:
            data = response.json()
            # Check if we have structured data
            for page in data.get('pages', []):
                page_data = page.get('data', {})
                if page_data.get('fields') or page_data.get('tables'):
                    return data
        time.sleep(interval)
    
    # Timeout, return last status
    response = requests.get(url, headers={'Authorization': auth_header}, timeout=30)
    return response.json()

def extract_table_data(doc_json):
    """Extract table data from the document JSON."""
    items = []
    for page in doc_json.get('pages', []):
        page_data = page.get('data', {})
        tables = page_data.get('tables', [])
        for table in tables:
            rows = table.get('rows', [])
            for row in rows:
                cells = row.get('cells', [])
                item = {}
                for cell in cells:
                    label = cell.get('label', '').strip()
                    text = cell.get('text', '').strip()
                    if label and text:
                        item[label] = text
                if item:
                    items.append(item)
    return items

def extract_fields(doc_json):
    """Extract field data from the document JSON."""
    fields = {}
    for page in doc_json.get('pages', []):
        page_data = page.get('data', {})
        for field in page_data.get('fields', []):
            label = field.get('label', '').strip()
            text = field.get('text', '').strip()
            if label and text:
                fields[label] = text
    return fields

def save_as_markdown(items, fields, image_path, output_dir):
    """Save results as a Markdown file."""
    md_lines = [
        f"# OCR Results - {Path(image_path).name}",
        "",
        f"**Workflow ID**: {WORKFLOW_ID}",
        f"**Processing Time**: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Extracted Fields",
    ]
    if fields:
        for k, v in fields.items():
            md_lines.append(f"- **{k}**: {v}")
    else:
        md_lines.append("_No fields detected_")
    md_lines.extend(["", "## Extracted Tables"])
    
    if items:
        # Assume all items have the same keys for a simple table
        if items:
            headers = list(items[0].keys())
            md_lines.append("| " + " | ".join(headers) + " |")
            md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
            for item in items:
                row = [item.get(h, "") for h in headers]
                md_lines.append("| " + " | ".join(row) + " |")
    else:
        md_lines.append("_No table data detected_")
    
    output_path = os.path.join(output_dir, f"{Path(image_path).stem}_result.md")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(md_lines))
    return output_path

def main():
    try:
        print(f"Uploading {IMAGE_PATH} to workflow {WORKFLOW_ID}...")
        upload_response = upload_document(WORKFLOW_ID, IMAGE_PATH)
        print(f"Upload response: {json.dumps(upload_response, indent=2)}")
        
        document_id = upload_response.get('document_id')
        if not document_id:
            print("Error: No document_id in response")
            return
        
        print(f"Waiting for processing to complete (document_id: {document_id})...")
        doc_json = wait_for_completion(WORKFLOW_ID, document_id)
        print(f"Processing complete. Response: {json.dumps(doc_json, indent=2)[:500]}...")
        
        # Save raw JSON
        json_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_raw.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(doc_json, f, indent=2, ensure_ascii=False)
        print(f"Raw JSON saved to {json_path}")
        
        # Extract data
        fields = extract_fields(doc_json)
        items = extract_table_data(doc_json)
        
        print(f"Extracted {len(fields)} fields and {len(items)} table rows")
        
        # Save as markdown
        md_path = save_as_markdown(items, fields, IMAGE_PATH, OUTPUT_DIR)
        print(f"Markdown report saved to {md_path}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
