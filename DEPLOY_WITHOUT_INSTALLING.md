# Deploy JRD Project Controls Hub without installing anything

The easiest route is Streamlit Community Cloud. You only need a browser.

## 1. Create a GitHub repository
- Go to GitHub.com and create a free account if you do not have one.
- Create a new repository, for example:
  `jrd-project-controls`
- Keep it PRIVATE if your company data is sensitive.

## 2. Upload these files from the browser
Upload the contents of this folder to the repository:
- `app.py`
- `requirements.txt`
- `README.md`
- `.streamlit/config.toml`

Do NOT upload live project data unless your company policy permits it.

## 3. Deploy on Streamlit Community Cloud
- Go to https://share.streamlit.io
- Sign in and connect your GitHub account.
- Click **Create app**
- Select your repository.
- Branch: `main`
- Main file path: `app.py`
- Click **Deploy**

After deployment you will receive a URL similar to:
`https://your-app-name.streamlit.app`

## 4. Use the app
Open the URL from any browser.
Upload the Primavera Excel export when needed.

The uploaded file is processed during the app session. Do not treat Community Cloud as an approved corporate storage location unless your company permits it.

## Recommended path for company use
For internal company deployment, later move the same app to:
- Microsoft Azure App Service / Container Apps, or
- your company's internal server

This is preferable for controlled access, SSO, corporate storage, and audit requirements.
