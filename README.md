# Huliot India Travel Expense Streamlit App

This version keeps an HTML-style interface inside Streamlit and saves entries live to Google Sheets when deployed.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud setup

1. Upload all files to GitHub.
2. Deploy the repository on Streamlit Cloud.
3. Create one Google Sheet.
4. Add the values from `.streamlit/secrets.toml.example` into Streamlit Cloud Secrets.
5. Share your Google Sheet with the service account `client_email`.

Do not upload your real `secrets.toml` to GitHub.

## How it works

- Save Entry writes data to Google Sheets.
- Saved Entries shows live saved data.
- Download Same Excel Format generates your attached Excel template and fills only entry cells.
