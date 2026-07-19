import requests
from requests.auth import HTTPBasicAuth
import json

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
MODEL_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"  # This is the workflow ID but also appears to work as model ID in v2
IMAGE_PATH = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"

auth = HTTPBasicAuth(API_KEY, '')

# Test v2 OCR Model endpoint
url = f"https://app.nanonets.com/api/v2/OCR/Model/{MODEL_ID}/LabelFile/"
files = {'file': ('unnamed (1).webp', open(IMAGE_PATH, 'rb'), 'image/webp')}
resp = requests.post(url, auth=auth, files=files, timeout=120)
print(f"POST OCR v2: {resp.status_code}")

try:
    result = resp.json()
    print(json.dumps(result, indent=2, ensure_ascii=False)[:2000])
except:
    print(resp.text[:2000])

# Save full response
with open("/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1)_v2_raw.json", "w", encoding="utf-8") as f:
    json.dump(resp.json(), f, indent=2, ensure_ascii=False)

