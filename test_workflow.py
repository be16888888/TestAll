import requests
import base64
import json

API_KEY = '09c019f3-8348-11f1-8b50-668e6851cbef'
auth = base64.b64encode(f'{API_KEY}:'.encode()).decode()
headers = {'Authorization': f'Basic {auth}'}

url = 'https://app.nanonets.com/api/v4/workflows/'
print(f"GET {url}")
response = requests.get(url, headers=headers, timeout=30)
print(f"Status: {response.status_code}")
print(f"Response: {response.text}")
try:
    print(f"JSON: {json.dumps(response.json(), indent=2)}")
except Exception as e:
    print(f"Not JSON: {e}")
