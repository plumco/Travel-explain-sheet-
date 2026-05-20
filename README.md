# Huliot Travel Expense Streamlit App

This version uses the uploaded HTML form design as the Streamlit interface style and keeps the original Excel workbook format for download.

## Files
- `app.py` - Streamlit app
- `Travelling Expenses Sheet.xlsx` - original Excel template
- `reference_html_form.html` - uploaded HTML reference
- `requirements.txt` - packages for Streamlit Cloud
- `.streamlit/secrets.toml.example` - example only, do not upload real secrets

## Streamlit Cloud setup
1. Upload all files to GitHub.
2. Deploy `app.py` on Streamlit Cloud.
3. Create one Google Sheet.
4. Share that Google Sheet with your service account email.
5. Add secrets in Streamlit Cloud settings.

If Google Sheets is not connected, the app saves to local CSV only for testing.
