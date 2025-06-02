# DMAX Web Report Generator

This app generates branded real estate deal reports from a spreadsheet and a Word template.

## 🔧 How to Use

1. Upload the DMAX spreadsheet (.xlsx)
2. Upload the DOCX report template with placeholders (e.g., {{PropertyAddress}})
3. Click 'Generate Report' to download a filled report

## 🔔 Optional: Zapier Notification

To enable webhook alerts to Zapier:
- Set the secret `zapier_webhook_url` in `.streamlit/secrets.toml` (locally)
- Or paste it in Streamlit Cloud under app settings → Secrets

## 🌐 Deploy on Streamlit Cloud

1. Upload all files to a GitHub repo
2. Go to [https://streamlit.io/cloud](https://streamlit.io/cloud)
3. Click 'New App', choose this repo, and set `dmax_web_report_app.py` as the main file