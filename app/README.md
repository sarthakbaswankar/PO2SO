# HTML to PDF API

A **FastAPI** service that accepts raw HTML from Microsoft Graph API email bodies,
strips Outlook-specific noise, and returns a pixel-faithful PDF.

---

## Project Structure

```
html-to-pdf-api/
├── app/
│   ├── __init__.py
│   ├── main.py          # FastAPI routes
│   └── converter.py     # HTML cleaning + WeasyPrint rendering
├── requirements.txt
├── render.yaml          # Render deployment config
├── Procfile
└── README.md
```

---

## Endpoints

| Method | Path       | Description                                  |
|--------|------------|----------------------------------------------|
| GET    | /          | Health check / welcome message               |
| GET    | /health    | Returns `{"status": "ok"}`                   |
| POST   | /convert   | Accepts JSON `{"html": "..."}`, returns PDF  |

---

## Local Development

### 1. Prerequisites

- Python 3.10+
- `pip`
- On Linux, WeasyPrint needs system fonts + Cairo:
  ```bash
  # Ubuntu/Debian
  sudo apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libcairo2

  # macOS (Homebrew)
  brew install pango cairo
  ```

### 2. Clone and install

```bash
git clone <your-repo-url>
cd html-to-pdf-api
pip install -r requirements.txt
```

### 3. Run locally

```bash
uvicorn app.main:app --reload --port 8000
```

API is now live at `http://localhost:8000`
Interactive docs at `http://localhost:8000/docs`

### 4. Test with curl

```bash
curl -X POST http://localhost:8000/convert \
  -H "Content-Type: application/json" \
  -d '{"html": "<html><body><p><b>Hello World</b></p></body></html>"}' \
  --output result.pdf
```

### 5. Test with the sample Graph API HTML

```bash
curl -X POST http://localhost:8000/convert \
  -H "Content-Type: application/json" \
  -d @- --output email.pdf <<EOF
{
  "html": "<html><body dir=\"ltr\"><div style=\"font-family:Calibri;font-size:12pt\">Testing the pdf - 2</div><table style=\"border-collapse:collapse\"><tr><td style=\"padding:0cm 5.4pt\"><img src=\"cid:abc\" width=\"167\" height=\"142\"></td><td style=\"padding:0cm 5.4pt\"><p><b>Harshit Kapse</b></p><p>Email: hkapse@acsesolutions.com</p><p>Mobile: +91-7000168807</p><p>Web: www.acsesolutions.com</p></td></tr></table></body></html>"
}
EOF
```

---

## Deploy on Render (Free Tier)

### Step 1 – Push to GitHub

```bash
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/<your-username>/html-to-pdf-api.git
git push -u origin main
```

### Step 2 – Create a Render account

Go to [https://render.com](https://render.com) and sign up (free, no credit card needed).

### Step 3 – New Web Service

1. In the Render dashboard click **"New +"** → **"Web Service"**
2. Connect your GitHub account and select the `html-to-pdf-api` repository
3. Render auto-detects `render.yaml` — confirm the settings:
   - **Environment**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
   - **Plan**: Free

### Step 4 – Add system-level dependencies (important for WeasyPrint)

WeasyPrint needs Pango/Cairo. In Render, add a build command that installs them:

In **render.yaml** update `buildCommand`:
```yaml
buildCommand: apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libcairo2 && pip install -r requirements.txt
```

Or in the Render dashboard under **Settings → Build Command** paste:
```
apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libcairo2 && pip install -r requirements.txt
```

### Step 5 – Deploy

Click **"Create Web Service"**. Render will build and deploy in ~2-3 minutes.
Your public URL will look like: `https://html-to-pdf-api.onrender.com`

### Step 6 – Test the live endpoint

```bash
curl -X POST https://html-to-pdf-api.onrender.com/convert \
  -H "Content-Type: application/json" \
  -d '{"html": "<html><body><p><b>Hello from Render!</b></p></body></html>"}' \
  --output result.pdf
```

> **Note**: The free Render tier spins down after 15 minutes of inactivity.
> The first request after sleep may take ~30 seconds to cold-start.

---

## How it works

1. **Input**: Raw HTML string from Microsoft Graph API (`/messages/{id}/$value` or `body.content`)
2. **Cleaning** (`converter.py`):
   - Strips broken Outlook `<style>` blocks
   - Removes noisy attributes (`data-outlook-trace`, `x_` IDs, `OWA*` IDs)
   - Replaces `cid:` embedded image references with a transparent placeholder
   - Keeps all inline styles (tables, font sizes, colours, line heights) intact
3. **Rendering**: WeasyPrint renders the clean HTML to PDF, preserving the layout exactly as it appeared in the email body
4. **Output**: Binary PDF streamed back as `application/pdf`
