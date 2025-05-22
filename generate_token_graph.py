# generate_token_graph.py
import msal
import os
from dotenv import load_dotenv
import atexit

load_dotenv() 

MS_GRAPH_CLIENT_ID = os.getenv('MS_GRAPH_CLIENT_ID')
MS_GRAPH_AUTHORITY = os.getenv('MS_GRAPH_AUTHORITY')
MS_GRAPH_SCOPES = ["User.Read", "Sites.Read.All", "Files.Read.All"] 

# --- Explicitly define cache file path relative to this script ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MS_GRAPH_TOKEN_CACHE_FILE = os.path.join(SCRIPT_DIR, "token_cache_ms_graph_chat.bin")
# --- End of path change ---

ms_graph_token_cache = msal.SerializableTokenCache()
if os.path.exists(MS_GRAPH_TOKEN_CACHE_FILE):
    try:
        ms_graph_token_cache.deserialize(open(MS_GRAPH_TOKEN_CACHE_FILE, "r").read())
        print(f"Token cache loaded from {MS_GRAPH_TOKEN_CACHE_FILE}")
    except Exception as e:
        print(f"Could not load token cache from '{MS_GRAPH_TOKEN_CACHE_FILE}', will create new: {e}")

def save_cache():
    if ms_graph_token_cache.has_state_changed:
        with open(MS_GRAPH_TOKEN_CACHE_FILE, "w") as cache_file:
            cache_file.write(ms_graph_token_cache.serialize())
        print(f"Token cache saved to {MS_GRAPH_TOKEN_CACHE_FILE}")
atexit.register(save_cache)

def main():
    if not MS_GRAPH_CLIENT_ID or not MS_GRAPH_AUTHORITY:
        print("Error: MS_GRAPH_CLIENT_ID or MS_GRAPH_AUTHORITY not set in .env file.")
        print("Please create/update your .env file with your Azure AD App registration details.")
        return

    print(f"Attempting to authenticate with scopes: {MS_GRAPH_SCOPES}")
    print(f"Using token cache file: {MS_GRAPH_TOKEN_CACHE_FILE}")


    app = msal.PublicClientApplication(
        MS_GRAPH_CLIENT_ID,
        authority=MS_GRAPH_AUTHORITY,
        token_cache=ms_graph_token_cache
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        print(f"Found account(s) in token cache for client ID {MS_GRAPH_CLIENT_ID}. Attempting to acquire token silently...")
        print(f"Accounts: {[acc['username'] for acc in accounts]}")
        chosen_account = accounts[0]
        result = app.acquire_token_silent(MS_GRAPH_SCOPES, account=chosen_account)

    if not result:
        print("No suitable token in cache or silent acquisition failed, initiating device flow...")
        flow = app.initiate_device_flow(scopes=MS_GRAPH_SCOPES)
        if "user_code" not in flow:
            print("Failed to create device flow. Error: " + flow.get("error_description", "Unknown error"))
            print(f"Full flow response: {flow}")
            return
        
        print(f"\nPlease authenticate with Microsoft Graph by navigating to: {flow['verification_uri']}")
        print(f"And entering the code: {flow['user_code']}")
        print(f"This code expires in {flow['expires_in'] // 60} minutes.")
        
        try:
            result = app.acquire_token_by_device_flow(flow) 
        except Exception as e:
            print(f"Error during device flow token acquisition: {e}")
            return

    if result and "access_token" in result:
        print("\nSuccessfully acquired Microsoft Graph access token.")
        print(f"Token will be cached in '{MS_GRAPH_TOKEN_CACHE_FILE}'. Your Flask app should now be able to use it.")
        
        user_info = result.get("id_token_claims", {})
        username_to_display = user_info.get('preferred_username')
        if not username_to_display and accounts:
             username_to_display = accounts[0].get('username')

        if username_to_display:
            print(f"Authenticated for user: {username_to_display}")
        else:
            print("Authenticated, but could not retrieve username from token claims.")
        
        # Explicitly call save_cache here to ensure it saves before script exits,
        # especially if no other state change happened after successful acquisition.
        ms_graph_token_cache.has_state_changed = True # Force save
        save_cache()

    elif result and "error" in result:
        print(f"\nError acquiring token: {result.get('error')}")
        print(f"Error description: {result.get('error_description')}")
        print(f"Full error result: {result}")
    else:
        print("\nCould not acquire token. Unknown error or flow cancelled by user.")

if __name__ == "__main__":
    main()