# generate_token_gmail.py
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import os

GMAIL_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send'
]
GMAIL_CREDENTIALS_FILE = 'credentials_gmail.json' # Ensure this file exists
GMAIL_TOKEN_FILE = 'token_gmail.json'

def main():
    creds = None
    # Check if token file exists and load it
    if os.path.exists(GMAIL_TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_FILE, GMAIL_SCOPES)
        except ValueError: # Handles malformed token.json
            print(f"Malformed '{GMAIL_TOKEN_FILE}', deleting to re-authenticate.")
            os.remove(GMAIL_TOKEN_FILE)
            creds = None
        except Exception as e:
            print(f"Error loading '{GMAIL_TOKEN_FILE}': {e}. Deleting to re-authenticate.")
            os.remove(GMAIL_TOKEN_FILE)
            creds = None


    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print("Refreshing existing Gmail token...")
                creds.refresh(GoogleAuthRequest())
                print("Token refreshed.")
            except Exception as e:
                print(f"Token refresh failed: {e}. Need to re-authenticate.")
                creds = None # Force re-authentication
                if os.path.exists(GMAIL_TOKEN_FILE): os.remove(GMAIL_TOKEN_FILE)
        
        if not creds: # If still no creds, or refresh failed
            if not os.path.exists(GMAIL_CREDENTIALS_FILE):
                print(f"Error: '{GMAIL_CREDENTIALS_FILE}' not found. Download it from Google Cloud Console.")
                return
            print(f"No valid token found or refresh failed. Initiating new authentication for scopes: {GMAIL_SCOPES}")
            flow = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_FILE, GMAIL_SCOPES)
            # The run_local_server will open a browser window for user authorization.
            # It starts a local web server to receive the authorization response.
            creds = flow.run_local_server(port=0) # port=0 finds an available port
        
        # Save the credentials for the next run
        with open(GMAIL_TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print(f"Token saved to '{GMAIL_TOKEN_FILE}' with scopes: {creds.scopes}")
    else:
        print(f"'{GMAIL_TOKEN_FILE}' is valid and up-to-date with required scopes.")
        print(f"Current scopes in token: {creds.scopes}")


if __name__ == '__main__':
    main()