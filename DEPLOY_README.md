# TikTok Live & Shop Dashboard Deployment

This app can be deployed as a public web app so other people can open a normal URL and upload their own TikTok exports.

## What Other People Will See

On a public deployment, the local-only buttons are hidden:

- `Load SHOPDATA + LIVEDATA` is only for this Mac.
- `Load CSVs from input_data/` is only for local testing.
- Public users should upload the four files on the page and click `Generate Dashboard`.

## Required Uploads

- Live Product export
- Shop Product export
- Live Performance export
- Shop Daily export

CSV, XLSX, and XLS are supported.

## Deploy On Render

1. In Render, create a new Web Service.
2. Select this GitHub repository.
3. Leave the root directory blank because `app.py` is in the repository root.
4. Use these settings:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Add environment variables:
   - `PUBLIC_DEPLOY=true`
   - `HOST=0.0.0.0`
6. Deploy.

After deployment, Render will give you a public `https://...` URL that you can share.

## Local Testing

Run:

```bash
python app.py
```

Then open:

```text
http://127.0.0.1:5050/
```

Do not open `templates/index.html` directly as a file. The upload buttons need the Python server.
