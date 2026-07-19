import requests
import base64
from pathlib import Path
import json

API_KEY = '09c019f3-8348-11f1-8b50-668e6851cbef'
WORKFLOW_ID = '46ae503e-bb1b-4c79-a5e7-86f4b812e1c7'
IMAGE_PATH = '/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp'

auth = base64.b64encode(f'{API_KEY}:'.encode()).decode()
headers = {'Authorization': f'Basic {auth}'}

url = f'https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}/documents/upload'

with open(IMAGE_PATH, 'rb') as f:
    files = {'file': (Path(IMAGE_PATH).name, f, 'image/webp')}
    data = {'async': 'false'}
    print(f"Uploading to {url}")
    response = requests.post(url, headers=headers, files=files, data=data, timeout=120)
    print(f"Status code: {response.status_code}")
    print(f"Response headers: {response.headers}")
    print(f"Response text: {response.text}")
    try:
        print(f"Response JSON: {response.json()}")
    except Exception as e:
        print(f"Failed to parse JSON: {e}")
