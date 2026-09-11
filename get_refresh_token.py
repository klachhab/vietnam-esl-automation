"""
get_refresh_token.py — Run this ONCE locally on your Mac.
==========================================================
It opens a browser for you to log in to your Google account,
then prints the refresh token you'll paste into GitHub Secrets.

Usage:
    pip install google-auth-oauthlib
    python get_refresh_token.py

You'll need client_secrets.json in the same folder (downloaded
from GCP Console → Credentials → OAuth 2.0 Client ID → Desktop app).
"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

def main():
    flow = InstalledAppFlow.from_client_secrets_file(
        "client_secrets.json",
        scopes=SCOPES,
    )
    # Opens your browser for login
    creds = flow.run_local_server(port=0)

    print("\n" + "="*60)
    print("COPY THESE 3 VALUES INTO GITHUB SECRETS:")
    print("="*60)
    print(f"\nGDRIVE_CLIENT_ID:\n  {creds.client_id}")
    print(f"\nGDRIVE_CLIENT_SECRET:\n  {creds.client_secret}")
    print(f"\nGDRIVE_REFRESH_TOKEN:\n  {creds.refresh_token}")
    print("\n" + "="*60)

    # Also save locally for reference (do NOT commit this file)
    with open("oauth_credentials.json", "w") as f:
        json.dump({
            "client_id": creds.client_id,
            "client_secret": creds.client_secret,
            "refresh_token": creds.refresh_token,
        }, f, indent=2)
    print("Also saved to oauth_credentials.json (DO NOT commit this file)")

if __name__ == "__main__":
    main()
