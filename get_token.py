import subprocess, sys, os, json

# Run from the script's own directory so client_secrets.json is found
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Auto-install required package if missing
try:
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("Installing google-auth-oauthlib...")
    subprocess.check_call([
        sys.executable, "-m", "pip", "install",
        "google-auth-oauthlib", "--break-system-packages"
    ])
    from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]
flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", SCOPES)
creds = flow.run_local_server(port=0)

with open("client_secrets.json") as f:
    secrets = json.load(f)

client_id     = secrets["installed"]["client_id"]
client_secret = secrets["installed"]["client_secret"]
refresh_token = creds.refresh_token

print("\n=== Copy these 3 values to Railway ===")
print(f"GOOGLE_CLIENT_ID     = {client_id}")
print(f"GOOGLE_CLIENT_SECRET = {client_secret}")
print(f"GOOGLE_REFRESH_TOKEN = {refresh_token}")
