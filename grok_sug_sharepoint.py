from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os

# Configuration
SITE_URL = "https://umair19hotmail.sharepoint.com/sites/EmailAIDrafterSharePoint"
CLIENT_ID = "26e8881d-753f-466a-b739-641c0eb82e04"
CLIENT_SECRET = "RmC8Q~9SpAqSsDWPgAgz-8T8tWtgwBrEOgftcaJw"
LIBRARY_NAME = "Documents"  # Double-check if this is "Shared Documents"
LOCAL_PATH = "./downloaded_files"

def connect_to_sharepoint():
    """Connect to SharePoint using client credentials."""
    try:
        credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
        ctx = ClientContext(SITE_URL).with_credentials(credentials)
        # Test connection by fetching web title
        web = ctx.web.get().execute_query()
        print(f"Connected to SharePoint site: {web.properties['Title']}")
        return ctx
    except Exception as e:
        raise Exception(f"Failed to connect to SharePoint: {str(e)}")

def download_files_from_library(ctx, library_name, local_path):
    """Download all files from the specified SharePoint library to local path."""
    try:
        # Ensure local directory exists
        if not os.path.exists(local_path):
            os.makedirs(local_path)
        
        # Get the document library
        library = ctx.web.lists.get_by_title(library_name)
        ctx.load(library)
        ctx.execute_query()
        print(f"Accessed library: {library.properties['Title']}")
        
        # Get all items in the library
        files = library.get_items().execute_query()
        if not files:
            print("No files found in the library.")
            return
        
        for item in files:
            # Get file properties
            file = item.file
            file_name = item.properties["FileLeafRef"]
            file_path = os.path.join(local_path, file_name)
            
            # Download file content
            with open(file_path, "wb") as local_file:
                file_content = file.get_content().execute_query()
                local_file.write(file_content)
            
            print(f"Downloaded: {file_name}")
    except Exception as e:
        raise Exception(f"Error downloading files: {str(e)}")

def main():
    try:
        # Connect to SharePoint
        ctx = connect_to_sharepoint()
        
        # Download files
        download_files_from_library(ctx, LIBRARY_NAME, LOCAL_PATH)
        print("All files downloaded successfully!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()