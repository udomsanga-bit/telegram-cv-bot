"""
Run this script ONCE locally to get your Google OAuth2 refresh token.

Steps:
1. Go to https://console.cloud.google.com/apis/credentials
2. Click "Create Credentials" → "OAuth 2.0 Client IDs"
3. Application type: Desktop app  →  Name: cv-bot-desktop  →  Create
4. Click the download icon (⬇) next to the new credential → save as client_secrets.json
   in this same folder (cv_bot/).
5. Run:  python get_token.py
6. A browser window opens → log in with udomsanga@gmail.com → Allow
7. Copy the three values printed below into Railway environment variables.
"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", SCOPES)
creds = flow.run_local_server(port=0)

with open("client_secrets.json") as f:
    secrets = json.load(f)

client_id     = secrets["installed"]["client_id"]
client_secret = secrets["installed"]["client_secret"]
refresh_token = creds.refresh_token

print("\n" + "="*60)
print("Add these 3 variables to Railway → worker → Variables:")
print("="*60)
print(f"GOOGLE_CLIENT_ID     = {client_id}")
print(f"GOOGLE_CLIENT_SECRET = {client_secret}")
print(f"GOOGLE_REFRESH_TOKEN = {refresh_token}")
print("="*60 + "\n")
