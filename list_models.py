import requests
from requests.auth import HTTPBasicAuth
import json

API_KEY = "09c019f3-8348-11f1-8b50-668e6851cbef"
auth = HTTPBasicAuth(API_KEY, '')

# List all models
url = "https://app.nanonets.com/api/v2/OCR/Model/"
resp = requests.get(url, auth=auth)
print(f"GET models: {resp.status_code}")
if resp.status_code == 200:
    models = resp.json()
    for model in models:
        print(f"  Model ID: {model.get('id')}, Name: {model.get('name')}, Status: {model.get('status')}, Type: {model.get('type')}")

