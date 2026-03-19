# 🏀 2026 Office March Madness — Web App

Mobile-friendly leaderboard + bracket viewer for your office pool.

## Files needed in this folder
```
march_madness_app/
├── app.py
├── requirements.txt
├── Procfile
├── README.md
├── templates/
│   └── index.html
└── 2026_Office_Bracket_UPGRADED.xlsx   ← copy this here from your project folder
```

---

## Run locally (test before deploying)

```bash
cd march_madness_app
pip install flask openpyxl gunicorn
python app.py
```
Open http://localhost:5000 in your browser.

---

## Deploy to Railway (FREE — recommended, 5 min setup)

Railway gives you a live public URL your friends can visit from their phones.

### Step 1 — Push to GitHub
1. Create a free account at github.com
2. Create a new repository called `march-madness-2026`
3. Upload all files from this folder INTO that repo (drag & drop works)
   - **Important:** include `2026_Office_Bracket_UPGRADED.xlsx` in the repo

### Step 2 — Deploy on Railway
1. Go to railway.app → Sign up free (use GitHub login)
2. Click **"New Project"** → **"Deploy from GitHub repo"**
3. Select your `march-madness-2026` repo
4. Railway auto-detects the Procfile and deploys automatically
5. Click **"Generate Domain"** under Settings → you get a public URL like:
   `https://march-madness-2026.up.railway.app`

### Step 3 — Share the link!
Send that URL to your friends. Works on any phone/browser, no login needed.

---

## Keeping results updated

Every time you run `update_bracket.py` on your local machine, it updates your
local Excel file. To push that update to the live website:

**Option A (easiest) — re-upload the Excel file to GitHub:**
1. Run `python update_bracket.py` locally
2. Go to your GitHub repo
3. Click on `2026_Office_Bracket_UPGRADED.xlsx` → Upload files → replace it
4. Railway auto-redeploys in ~30 seconds

**Option B — run update_bracket.py on Railway directly:**
In Railway dashboard → your project → Add a Cron Job:
- Command: `python update_bracket.py`
- Schedule: `*/30 * * * *`  (every 30 minutes)
This runs automatically all tournament long — totally hands-off!

---

## Alternative: Render.com (also free)

1. Same GitHub setup as above
2. Go to render.com → New → Web Service
3. Connect your GitHub repo
4. Set Start Command to: `gunicorn app:app`
5. Click Deploy → get your public URL

---

## Local network sharing (quickest option)

If you just want friends on the same WiFi to see it:
```bash
python app.py
```
Then share your local IP address (e.g. http://192.168.1.5:5000).
Find your IP: run `ipconfig` (Windows) or `ifconfig` (Mac).
