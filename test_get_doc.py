import requests
from requests.auth import HTTPBasicAuth
import json
import time

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
WORKFLOW_ID = "46ae503e-bb1b-4c79-a5e7-86f4b812e1c7"
DOCUMENT_ID = "b2d2f0f1-eb95-4692-a046-cca27c87abf8"

auth = HTTPBasicAuth(API_KEY, '')

# Get document
url = f"https://app.nanonets.com/api/v4/workflows/{WORKFLOW_ID}/documents/{DOCUMENT_ID}"
resp = requests.get(url, auth=auth)
print(f"GET document: {resp.status_code}")
print(resp.text[:1000])

