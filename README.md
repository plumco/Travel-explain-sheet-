# Huliot India Travel Expense Entry App

This Streamlit app saves travel expense entries directly into Google Sheets and downloads the same Excel template format.

## Files
- `app.py` - Streamlit app
- `requirements.txt` - Python packages
- `Travelling Expenses Sheet.xlsx` - original Excel template
- `.streamlit/secrets.toml.example` - example secrets format

## Local Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Google Sheets Setup
1. Create one Google Sheet.
2. Copy the Sheet ID from the URL.
3. Create Google Cloud service account and JSON key.
4. Share the Google Sheet with the service account email.
5. In Streamlit Cloud, add secrets using `.streamlit/secrets.toml.example` format.
6. Deploy from GitHub.

## Important
Do not upload the real `secrets.toml` file to GitHub. Add secrets in Streamlit Cloud settings.
