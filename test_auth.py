import requests
from requests.auth import HTTPBasicAuth
import json

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
WORKFLOW_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"
IMAGE_PATH = "/mnt/e/DiskCUse/HFDownloads/OCRUse02/unnamed (1).webp"

# Test with HTTPBasicAuth
auth = HTTPBasicAuth(API_KEY, '')

# Test 1: Get workflows
url = "https://app.nanonets.com/api/v4/workflows/"
resp = requests.get(url, auth=auth)
print(f"GET workflows: {resp.status_code}")
print(resp.text[:200])

# Test 2: Get specific workflow
url = f"https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}"
resp = requests.get(url, auth=auth)
print(f"\nGET workflow: {resp.status_code}")
print(resp.text[:200])

# Test 3: Upload document
url = f"https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}/documents/"
files = {'file': ('unnamed (1).webp', open(IMAGE_PATH, 'rb'), 'image/webp')}
data = {'async': 'false'}
resp = requests.post(url, auth=auth, files=files, data=data, timeout=120)
print(f"\nPOST upload: {resp.status_code}")
print(resp.text[:500])

