# Deployment Guide — Google Drive + Railway

This guide sets up two things:
1. **Google Drive** — CVs and the Excel log are saved to a shared Google Drive folder (free, no subscription)
2. **Railway** — the bot runs 24/7 on a cloud server (no computer needed, free tier available)

---

## Part 1 — Connect Google Drive (~10 minutes)

### Step 1 — Create a Google Cloud Project
1. Go to: https://console.cloud.google.com
2. Sign in with any Google account (personal Gmail is fine)
3. Click the project dropdown at the top → **"New Project"**
4. Name it `cv-bot` → click **Create**

### Step 2 — Enable Google Drive API
1. In the search bar type **"Google Drive API"** → click it
2. Click **"Enable"**

### Step 3 — Create a Service Account
1. Go to **APIs & Services → Credentials** in the left menu
2. Click **"+ Create Credentials"** → **"Service account"**
3. Fill in:
   - Service account name: `cv-bot`
   - Click **Create and Continue** → **Done** (skip optional steps)
4. Click the service account you just created
5. Go to the **"Keys"** tab → **"Add Key"** → **"Create new key"** → **JSON** → **Create**
6. A `.json` file downloads automatically — **keep this safe**

### Step 4 — Encode the key for Railway
Railway can't read a file directly, so we encode it as text.

Open Terminal on your Mac and run (replace the filename):
```bash
base64 -i ~/Downloads/cv-bot-xxxx.json | tr -d '\n'
```
Copy the long string of text that appears — this is your `GOOGLE_SERVICE_ACCOUNT_JSON`.

### Step 5 — Create a shared Google Drive folder
1. Go to https://drive.google.com
2. Click **"+ New"** → **"New folder"** → name it `CVs`
3. Right-click the folder → **"Share"**
4. Open the `.json` file you downloaded and find the `"client_email"` field.
   It looks like: `cv-bot@your-project.iam.gserviceaccount.com`
5. Paste that email into the Share box → set role to **"Editor"** → click **Send**

### Step 6 — Get the folder ID
1. Open the `CVs` folder in Google Drive
2. Look at the URL: `https://drive.google.com/drive/folders/`**`1ABC123xyz`**
3. Copy the part after `/folders/` — this is your `GOOGLE_DRIVE_FOLDER_ID`

---

## Part 2 — Deploy to Railway (10 minutes)

Railway runs your bot permanently on a cloud server — no computer needed.

### Step 1 — Push code to GitHub
1. Create a free account at https://github.com
2. Create a **private** repository called `cv-bot`
3. Upload all files in the `cv_bot` folder to that repo

   In Terminal:
   ```bash
   cd "/path/to/cv_bot"
   git init
   git add .
   git commit -m "initial"
   git remote add origin https://github.com/YOUR_USERNAME/cv-bot.git
   git push -u origin main
   ```

### Step 2 — Create a Railway project
1. Go to https://railway.app → sign up free
2. Click **"New Project"** → **"Deploy from GitHub repo"**
3. Connect GitHub and select `cv-bot`
4. Railway detects Python automatically and starts deploying

### Step 3 — Add Environment Variables
1. Click your service → **"Variables"** tab
2. Add these three variables:

| Variable | Value |
|---|---|
| `BOT_TOKEN` | Your Telegram bot token |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The base64 string from Part 1 Step 4 |
| `GOOGLE_DRIVE_FOLDER_ID` | The folder ID from Part 1 Step 6 |

3. Railway auto-redeploys with the new variables

### Step 4 — Confirm it's working
Click **"Deployments"** → latest deployment → **"View Logs"**

You should see:
```
Webhook cleared. Polling ready.
Bot online: @yourbotname
```

Send a test CV on Telegram — the file should appear in your `CVs` Google Drive folder within seconds, and `submissions.xlsx` will be created/updated there too.

---

## What ends up in Google Drive

```
My Drive/
└── CVs/                          ← the shared folder you created
    ├── submissions.xlsx           ← live log, updated on every submission
    ├── 20240617_143022_john_cv.pdf
    ├── 20240617_151055_sara_resume.docx
    └── ...
```

You can share the `CVs` folder with your whole HR team — they'll see new CVs appear in real time.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `Google Drive upload failed` | Check the service account email was added as **Editor** to the folder |
| Files appear but Excel doesn't update | Check Railway logs for permission errors |
| Bot not responding | Verify `BOT_TOKEN` is correct in Railway variables |
| `invalid_grant` error | Re-download the service account JSON key and re-encode it |
