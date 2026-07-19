import requests
import time
import json
import os
from pathlib import Path
from PIL import Image

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
WORKFLOW_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"
IMAGE_PATH = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"
OUTPUT_DIR = "/home/newuser/TestAll/ocr_results"

os.makedirs(OUTPUT_DIR, exist_ok=True)

def convert_to_jpeg(image_path):
    """Convert any non-JPEG image to .jpeg format for Nanonets compatibility."""
    path = Path(image_path)
    suffix = path.suffix.lower()
    if suffix in ('.jpeg', '.jpg'):
        return image_path
    
    jpeg_path = path.with_suffix('.jpeg')
    
    with Image.open(image_path) as img:
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        
        img.save(jpeg_path, 'JPEG', quality=95)
    
    print(f"Converted: {image_path} -> {jpeg_path}")
    return str(jpeg_path)

def extract_html_content(json_data):
    """從 API 返回的 JSON 中提取 markdown.content (實際上是 HTML 表格)"""
    try:
        # 111.py 使用的路徑: result -> markdown -> content
        content = json_data.get("result", {}).get("markdown", {}).get("content")
        if content is None:
            content = json_data.get("markdown", {}).get("content")
        if content is None:
            # 有時 result 是列表
            if isinstance(json_data.get("result"), list) and len(json_data["result"]) > 0:
                first = json_data["result"][0]
                content = first.get("markdown", {}).get("content")
        if content is None:
            raise ValueError("无法从 JSON 中找到 markdown content")
        return content
    except Exception as e:
        raise Exception(f"解析 JSON 失败: {e}\n原始响应: {json_data}")

def save_html(content, image_path, output_dir):
    """保存 HTML 內容到 .md 文件"""
    md_path = os.path.join(output_dir, f"{Path(image_path).stem}_result.md")
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# OCR Results - {Path(image_path).name}\n\n")
        f.write(f"**Processing Time**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(content)
    return md_path

def main():
    try:
        # Convert image to JPEG if needed
        processed_image_path = convert_to_jpeg(IMAGE_PATH)
        
        print(f"Uploading {processed_image_path}...")
        
        # 參考 111.py: 使用 extraction-api.nanonets.com v1 端點
        url = "https://extraction-api.nanonets.com/api/v1/extract/sync"
        headers = {"Authorization": f"Bearer {API_KEY}"}
        
        with open(processed_image_path, 'rb') as f:
            files = {'file': (Path(processed_image_path).name, f)}
            data = {'output_format': 'markdown'}
            resp = requests.post(url, headers=headers, files=files, data=data)
        
        resp.raise_for_status()
        json_data = resp.json()
        print(f"Upload response: {json.dumps(json_data, indent=2)}")
        
        # Save raw JSON
        json_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_raw.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        print(f"Raw JSON saved to {json_path}")
        
        # Extract HTML content (markdown.content)
        try:
            html_content = extract_html_content(json_data)
            md_path = save_html(html_content, IMAGE_PATH, OUTPUT_DIR)
            print(f"HTML/Markdown report saved to {md_path}")
        except Exception as e:
            print(f"無法提取 markdown content: {e}")
            # Fallback
            md_path = os.path.join(OUTPUT_DIR, f"{Path(IMAGE_PATH).stem}_result.md")
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(f"# OCR Results - {Path(IMAGE_PATH).name}\n\n")
                f.write(f"**Processing Time**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write("## Raw JSON Response\n```json\n")
                f.write(json.dumps(json_data, indent=2, ensure_ascii=False))
                f.write("\n```\n")
            print(f"Fallback markdown saved to {md_path}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Clean up temporary .jpeg file if it was created
        if 'processed_image_path' in locals() and processed_image_path != IMAGE_PATH:
            try:
                os.remove(processed_image_path)
                print(f"Cleaned up temporary file: {processed_image_path}")
            except:
                pass

if __name__ == "__main__":
    main()