# app.py
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import platform
import csv
import json
import requests as ollama_requests
import random
import re
import base64
from email.mime.text import MIMEText
from datetime import datetime
import time
from dotenv import load_dotenv # For .env file

# Google API Client Libraries
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# MSAL for Microsoft Graph (SharePoint)
import msal
import atexit

# Document parsing libraries
from docx import Document as DocxDocument # python-docx
import fitz  # PyMuPDF

# --- Load Environment Variables ---
load_dotenv()

# --- Flask App Setup ---
app = Flask(__name__)
CORS(app, supports_credentials=True)
app.secret_key = os.urandom(32)

# --- Configuration ---
OLLAMA_API_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "deepseek-r1:14b") # Use env var or default
NUM_STYLE_EXAMPLES = int(os.getenv("NUM_STYLE_EXAMPLES", 3))
USER_NAME = os.getenv("USER_NAME", "Khurram")

# Gmail Config
GMAIL_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.metadata'
]
GMAIL_CREDENTIALS_FILE = 'credentials_gmail.json'
GMAIL_TOKEN_FILE = 'token_gmail.json'

# Outlook Desktop COM Config (Windows Only)
if platform.system() == "Windows":
    import win32com.client
    import pythoncom
    class COMScope: # Defined as before
        def __enter__(self):
            app.logger.debug("Initializing COM for thread...")
            pythoncom.CoInitializeEx(0)
            return self
        def __exit__(self, exc_type, exc_val, exc_tb):
            app.logger.debug("Uninitializing COM for thread...")
            pythoncom.CoUninitialize()
else:
    class COMScope:
        def __enter__(self): return self
        def __exit__(self, exc_type, exc_val, exc_tb): pass

# Microsoft Graph (SharePoint) Config
MS_GRAPH_CLIENT_ID = os.getenv('MS_GRAPH_CLIENT_ID')
MS_GRAPH_AUTHORITY = os.getenv('MS_GRAPH_AUTHORITY') # e.g. https://login.microsoftonline.com/YOUR_TENANT_ID
MS_GRAPH_SCOPES = ["User.Read", "Sites.Read.All", "Files.Read.All"] # Add Mail scopes if reusing for Outlook email
MS_GRAPH_TOKEN_CACHE_FILE = "token_cache_graph.bin"

SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME", "DefaultSiteName") # Specify your default site
SHAREPOINT_DEFAULT_DRIVE_NAME = os.getenv("SHAREPOINT_DEFAULT_DRIVE_NAME", "Documents")

# --- MSAL Token Cache for Graph API ---
ms_graph_token_cache = msal.SerializableTokenCache()
if os.path.exists(MS_GRAPH_TOKEN_CACHE_FILE):
    try:
        ms_graph_token_cache.deserialize(open(MS_GRAPH_TOKEN_CACHE_FILE, "r").read())
        app.logger.info(f"MS Graph token cache loaded from {MS_GRAPH_TOKEN_CACHE_FILE}")
    except Exception as e:
        app.logger.error(f"Error loading MS Graph token cache: {e}. A new one may be created.")

def save_ms_graph_cache():
    if ms_graph_token_cache.has_state_changed:
        with open(MS_GRAPH_TOKEN_CACHE_FILE, "w") as cache_file:
            cache_file.write(ms_graph_token_cache.serialize())
        app.logger.info(f"MS Graph token cache saved to {MS_GRAPH_TOKEN_CACHE_FILE}")
atexit.register(save_ms_graph_cache)


# --- Gmail Functions (get_gmail_service, fetch_gmail_emails_internal, parse_gmail_body, get_gmail_email_details_internal, send_gmail_reply_internal) ---
# These should be THE EXACT SAME as the versions that were "working good" for you previously.
# For brevity, I'm not re-pasting them here, but ensure you use your last working versions.
# Make sure they use app.logger.
def get_gmail_service(interactive_auth_if_needed=False):
    # ... (Your working version from the previous response) ...
    # (Ensure it correctly uses GMAIL_TOKEN_FILE, GMAIL_CREDENTIALS_FILE, GMAIL_SCOPES and app.logger)
    creds = None
    script_dir = os.path.dirname(__file__) 
    token_path = os.path.join(script_dir, GMAIL_TOKEN_FILE)
    creds_path = os.path.join(script_dir, GMAIL_CREDENTIALS_FILE)

    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, GMAIL_SCOPES)
        except ValueError: # Handles malformed token.json
            app.logger.warning(f"Malformed Gmail token file at {token_path}. Deleting.")
            if os.path.exists(token_path): os.remove(token_path)
            creds = None
        except Exception as e:
            app.logger.warning(f"Error loading Gmail token from {token_path}: {e}. Deleting.")
            if os.path.exists(token_path): os.remove(token_path)
            creds = None
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                app.logger.info("Attempting to refresh Gmail token...")
                creds.refresh(GoogleAuthRequest())
                app.logger.info("Gmail token refreshed successfully.")
            except Exception as e:
                app.logger.error(f"Error refreshing Gmail token: {e}. Will attempt re-authentication if interactive.")
                creds = None 
                if os.path.exists(token_path): os.remove(token_path)
        
        if not creds and interactive_auth_if_needed:
            if not os.path.exists(creds_path):
                app.logger.error(f"'{creds_path}' not found for Gmail auth.")
                return None
            try:
                flow = InstalledAppFlow.from_client_secrets_file(creds_path, GMAIL_SCOPES)
                app.logger.info("Initiating new Gmail OAuth flow (will open browser on server)...")
                creds = flow.run_local_server(port=0) 
                app.logger.info("Gmail OAuth flow completed.")
            except Exception as e:
                app.logger.error(f"Failed to run Gmail interactive auth flow: {e}", exc_info=True)
                return None
        elif not creds:
            app.logger.warning(f"No valid Gmail credentials. Token file: '{token_path}'. Required scopes: {GMAIL_SCOPES}")
            return None

    if creds: 
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())
            
    try:
        service = build('gmail', 'v1', credentials=creds)
        app.logger.info("Gmail service built successfully.")
        return service
    except Exception as e:
        app.logger.error(f"Failed to build Gmail service: {e}", exc_info=True)
    return None

def parse_gmail_body(payload, message_id="UnknownMsg"):
    # ... (Your working version from the previous response with extensive logging) ...
    if not payload:
        app.logger.warning(f"MsgID {message_id}: No payload provided to parse_gmail_body")
        return ""
    mime_type = payload.get('mimeType', '')
    app.logger.debug(f"MsgID {message_id}: parse_gmail_body - Processing payload/part with mimeType: '{mime_type}', Filename: {payload.get('filename')}")
    body_content_found = ""
    if 'parts' in payload:
        app.logger.debug(f"MsgID {message_id}: multipart message with {len(payload['parts'])} parts.")
        plain_text_body_from_parts = None
        html_body_from_parts = None
        for i, part in enumerate(payload['parts']):
            part_mime_type = part.get('mimeType', '')
            app.logger.debug(f"MsgID {message_id}: Checking sub-part {i}, mimeType: '{part_mime_type}', Filename: {part.get('filename')}")
            if part_mime_type == 'text/plain':
                part_body_data = part.get('body', {}).get('data')
                if part_body_data:
                    try:
                        decoded_data = base64.urlsafe_b64decode(part_body_data).decode('utf-8', 'replace')
                        if plain_text_body_from_parts is None: plain_text_body_from_parts = decoded_data
                        app.logger.debug(f"MsgID {message_id}: Decoded text/plain data (len {len(decoded_data)}).")
                        break # Found ideal plain text
                    except Exception as e: app.logger.error(f"MsgID {message_id}: Error decoding text/plain sub-part {i} data: {e}")
            elif part_mime_type == 'text/html':
                part_body_data = part.get('body', {}).get('data')
                if part_body_data:
                    try:
                        decoded_data = base64.urlsafe_b64decode(part_body_data).decode('utf-8', 'replace')
                        if html_body_from_parts is None: html_body_from_parts = decoded_data
                        app.logger.debug(f"MsgID {message_id}: Decoded text/html data (len {len(decoded_data)}).")
                    except Exception as e: app.logger.error(f"MsgID {message_id}: Error decoding text/html sub-part {i} data: {e}")
            elif part_mime_type.startswith('multipart/'):
                app.logger.debug(f"MsgID {message_id}: Recursing into nested multipart sub-part {i}: {part_mime_type}")
                nested_body = parse_gmail_body(part, message_id=f"{message_id}-sub{i}")
                if nested_body:
                    app.logger.debug(f"MsgID {message_id}: Found body in deeply nested multipart (sub-part {i}): '{nested_body[:50]}...'")
                    return nested_body.strip() 
        if plain_text_body_from_parts is not None:
            body_content_found = plain_text_body_from_parts
        elif html_body_from_parts is not None:
            temp_body = re.sub(r'<style([\S\s]*?)</style>','', html_body_from_parts, flags=re.DOTALL|re.IGNORECASE)
            temp_body = re.sub(r'<script([\S\s]*?)</script>','', temp_body, flags=re.DOTALL|re.IGNORECASE)
            temp_body = re.sub(r'<head([\S\s]*?)</head>','', temp_body, flags=re.DOTALL|re.IGNORECASE)
            temp_body = re.sub(r'<p[^>]*>', '\n', temp_body); temp_body = re.sub(r'<br\s*/?>', '\n', temp_body)
            body_content_found = " ".join(re.sub('<[^<]+?>', ' ', temp_body).split()).strip()
    elif 'body' in payload and 'data' in payload['body']: 
        app.logger.debug(f"MsgID {message_id}: Processing single part payload with body.data, mimeType: {mime_type}")
        data = payload['body']['data']
        try:
            body_data = base64.urlsafe_b64decode(data).decode('utf-8', 'replace')
            if mime_type == 'text/plain': body_content_found = body_data
            elif mime_type == 'text/html':
                temp_body = re.sub(r'<style([\S\s]*?)</style>','', body_data, flags=re.DOTALL|re.IGNORECASE)
                temp_body = re.sub(r'<script([\S\s]*?)</script>','', temp_body, flags=re.DOTALL|re.IGNORECASE)
                temp_body = re.sub(r'<head([\S\s]*?)</head>','', temp_body, flags=re.DOTALL|re.IGNORECASE)
                temp_body = re.sub(r'<p[^>]*>', '\n', temp_body); temp_body = re.sub(r'<br\s*/?>', '\n', temp_body)
                body_content_found = " ".join(re.sub('<[^<]+?>', ' ', temp_body).split()).strip()
        except Exception as e: app.logger.error(f"MsgID {message_id}: Error decoding single part data: {e}")
    else: app.logger.warning(f"MsgID {message_id}: No 'parts' and no direct 'body.data' found. Keys: {list(payload.keys())}. Filename: {payload.get('filename')}")
    return body_content_found.strip()

def get_gmail_email_details_internal(service, message_id):
    # ... (Your working version from the previous response, using the updated parse_gmail_body)
    if not service: app.logger.error(f"Gmail service N/A for details of {message_id}"); return None
    try:
        app.logger.info(f"Fetching FULL Gmail details for message ID: {message_id}")
        msg = service.users().messages().get(userId='me', id=message_id, format='full').execute()
        payload = msg.get('payload', {})
        if not payload: app.logger.error(f"MsgID {message_id}: No payload in fetched message."); return None
        headers = payload.get('headers', [])
        email_details = { "id": msg.get('id'), "platform": "gmail", "body": "", "from": "", "to": "", "cc": "", "subject": "", "threadId": msg.get('threadId'), "message_id_header": "", "references_header": "", "in_reply_to_header": "" }
        for header in headers:
            name = header['name'].lower()
            if name == 'subject': email_details['subject'] = header['value']
            elif name == 'from': email_details['from'] = header['value']
            elif name == 'to': email_details['to'] = header['value']
            elif name == 'cc': email_details['cc'] = header['value']
            elif name == 'message-id': email_details['message_id_header'] = header['value']
            elif name == 'references': email_details['references_header'] = header['value']
            elif name == 'in-reply-to': email_details['in_reply_to_header'] = header['value']
        email_details['body'] = parse_gmail_body(payload, message_id=message_id)
        if not email_details['body']: email_details['body'] = msg.get('snippet', '[Body not extracted]')
        app.logger.info(f"Processed Gmail details for ID: {message_id}. Body len: {len(email_details['body'])}.")
        return email_details
    except HttpError as error:
        err_c = error.content.decode() if hasattr(error, 'content') and error.content else ""
        app.logger.error(f"Gmail API HttpError getting details for {message_id}: Status {error.resp.status if hasattr(error,'resp') else 'N/A'}, Reason {error._get_reason()}, Content: {err_c}", exc_info=True)
    except Exception as e: app.logger.error(f"General error getting Gmail details for {message_id}: {e}", exc_info=True)
    return None

def fetch_gmail_emails_internal(service, folder_label, count, for_style=False):
    # ... (Your working version from the previous response, ensure for_style=True uses full parse for body)
    emails_list = []
    if not service: return emails_list
    try:
        q_str = "category:primary" if folder_label == 'SENT' and for_style else None
        results = service.users().messages().list(userId='me', labelIds=[folder_label], maxResults=count, q=q_str).execute()
        messages_info = results.get('messages', [])
        for msg_info in messages_info:
            if for_style: 
                try:
                    msg = service.users().messages().get(userId='me', id=msg_info['id'], format='full').execute() # Fetch full for body
                    body_content = parse_gmail_body(msg.get('payload',{}), message_id=f"{msg_info['id']}-style")
                    if body_content and len(body_content) > 30 : emails_list.append({"body": body_content})
                except Exception as e_style: app.logger.warning(f"Error processing Gmail msg {msg_info['id']} for style: {e_style}")
            else: 
                msg = service.users().messages().get(userId='me', id=msg_info['id'], format='metadata', metadataHeaders=['Subject', 'From', 'Date', 'To', 'Cc', 'Message-ID', 'References', 'In-Reply-To']).execute()
                payload = msg.get('payload', {}); headers = payload.get('headers', [])
                email_data = { "id": msg.get('id'), "threadId": msg.get('threadId'), "snippet": msg.get('snippet', '').strip(), "platform": "gmail", "subject": "", "from": "", "date": "", "to": "", "cc": "", "message_id_header": "", "references_header": "", "in_reply_to_header": "" }
                for header in headers:
                    name = header['name'].lower()
                    if name == 'subject': email_data['subject'] = header['value']
                    elif name == 'from': email_data['from'] = header['value']
                    elif name == 'date': email_data['date'] = header['value']
                    elif name == 'to': email_data['to'] = header['value']
                    elif name == 'cc': email_data['cc'] = header['value']
                    elif name == 'message-id': email_data['message_id_header'] = header['value']
                    elif name == 'references': email_data['references_header'] = header['value']
                    elif name == 'in-reply-to': email_data['in_reply_to_header'] = header['value']
                emails_list.append(email_data)
    except Exception as e: app.logger.error(f"Error fetching Gmail {folder_label}: {e}", exc_info=True)
    return emails_list

def send_gmail_reply_internal(service, to_recipients, subject, body, thread_id, in_reply_to_header=None, references_header=None):
    # ... (Your working version from the previous response) ...
    if not service: return False, "Gmail service not available."
    try:
        message = MIMEText(body); message['to'] = to_recipients; message['subject'] = subject
        if in_reply_to_header: message['In-Reply-To'] = in_reply_to_header
        if references_header: message['References'] = references_header
        elif in_reply_to_header : message['References'] = in_reply_to_header
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        send_payload = {'raw': raw_message}; 
        if thread_id: send_payload['threadId'] = thread_id
        sent_message = service.users().messages().send(userId='me', body=send_payload).execute()
        app.logger.info(f"Gmail reply sent. ID: {sent_message.get('id')}")
        return True, sent_message.get('id')
    except Exception as e: app.logger.error(f"Error sending Gmail reply: {e}", exc_info=True)
    return False, str(e)

# --- Outlook Desktop Functions (fetch_outlook_emails_internal, get_outlook_email_details_internal, send_outlook_reply_internal) ---
# These should be THE EXACT SAME as the versions that were "working good" for you previously.
# Ensure they use COMScope and app.logger.
# For brevity, I'm not re-pasting them here, but ensure you use your last working versions.
def fetch_outlook_emails_internal(mapi_folder_constant, count, for_style=False):
    # ... (Your working version from previous app.py, using COMScope)
    emails_list = []
    if platform.system() != "Windows": app.logger.info("Skipping Outlook fetch: Not on Windows."); return emails_list
    try:
        with COMScope(): 
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            if not outlook_app: app.logger.error("Failed to dispatch Outlook within COMScope."); return emails_list
            namespace = outlook_app.GetNamespace("MAPI"); folder = namespace.GetDefaultFolder(mapi_folder_constant) 
            app.logger.info(f"Accessing Outlook folder: {folder.Name} (Const: {mapi_folder_constant})")
            messages = folder.Items; messages.Sort("[ReceivedTime]", True)
            processed_count = 0; items_to_check = min(messages.Count if messages.Count else 0, count * 5 + 20) 
            for i in range(1, items_to_check + 1): 
                if processed_count >= count: break
                try:
                    message = messages.Item(i)
                    if message is None or message.Class != 43: continue
                    if for_style:
                        body_content = message.Body
                        if body_content and len(body_content) > 30: emails_list.append({"body": body_content[:1500]})
                    else:
                        sender_email_val = ""; 
                        try:
                            if message.SenderEmailType == "EX": sender_email_val = message.Sender.GetExchangeUser().PrimarySmtpAddress
                            else: sender_email_val = message.SenderEmailAddress
                        except: sender_email_val = message.SenderEmailAddress if hasattr(message, 'SenderEmailAddress') and message.SenderEmailAddress else (message.SenderName if hasattr(message, 'SenderName') else "Unknown")
                        try: date_obj = message.ReceivedTime 
                        except: date_obj = None
                        date_str = date_obj.strftime("%a, %d %b %Y %H:%M:%S %z") if date_obj else datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
                        prop_accessor = message.PropertyAccessor
                        msg_id_h,in_reply_to_h,refs_h = "","",""
                        try: msg_id_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                        except:pass
                        try: in_reply_to_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1042001F")
                        except:pass
                        try: refs_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1039001F")
                        except:pass
                        email_data = {"id": message.EntryID, "threadId": message.ConversationID, "subject": message.Subject or "(No Subject)", "from": f"{message.SenderName or 'Unknown Sender'} <{sender_email_val or 'N/A'}>", "to": message.To or "", "date": date_str, "snippet": (message.Body[:150].replace("\r\n", " ").strip() if message.Body else ""), "platform": "outlook", "message_id_header": msg_id_h, "in_reply_to_header": in_reply_to_h, "references_header": refs_h}
                        emails_list.append(email_data)
                    processed_count += 1
                except Exception as e_item: app.logger.warning(f"Could not process Outlook item {i} in {folder.Name}: {e_item}", exc_info=False)
    except Exception as e: app.logger.error(f"Error fetching Outlook emails (folder {mapi_folder_constant}): {e}", exc_info=True)
    return emails_list

def get_outlook_email_details_internal(entry_id):
    # ... (Your working version from previous app.py, using COMScope)
    if platform.system() != "Windows": return None
    try:
        with COMScope():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            if not outlook_app: return None
            namespace = outlook_app.GetNamespace("MAPI"); mail_item = namespace.GetItemFromID(entry_id)
            if mail_item and mail_item.Class == 43:
                sender_email_val = ""; 
                try:
                    if mail_item.SenderEmailType == "EX": sender_email_val = mail_item.Sender.GetExchangeUser().PrimarySmtpAddress
                    else: sender_email_val = mail_item.SenderEmailAddress
                except: sender_email_val = mail_item.SenderEmailAddress if hasattr(mail_item, 'SenderEmailAddress') and mail_item.SenderEmailAddress else (mail_item.SenderName if hasattr(mail_item, 'SenderName') else "Unknown")
                prop_accessor = mail_item.PropertyAccessor
                msg_id_h,in_reply_to_h,refs_h = "","",""
                try: msg_id_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                except:pass
                try: in_reply_to_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1042001F")
                except:pass
                try: refs_h=prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1039001F")
                except:pass
                try: date_str = mail_item.SentOn.strftime("%a, %d %b %Y %H:%M:%S %z") if mail_item.SentOn else (mail_item.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S %z") if mail_item.ReceivedTime else "")
                except: date_str = ""
                details = {"id": mail_item.EntryID, "platform": "outlook", "subject": mail_item.Subject or "(No Subject)", "from": f"{mail_item.SenderName or 'Unknown Sender'} <{sender_email_val or 'N/A'}>", "to": mail_item.To or "", "cc": mail_item.CC or "", "date": date_str, "body": mail_item.Body or "", "html_body": mail_item.HTMLBody or "", "threadId": mail_item.ConversationID, "message_id_header": msg_id_h, "references_header": refs_h, "in_reply_to_header": in_reply_to_h}
                return details
    except Exception as e: app.logger.error(f"Error getting Outlook details for EntryID {entry_id}: {e}", exc_info=True)
    return None

def send_outlook_reply_internal(original_entry_id, to_recipients_str, subject, body):
    # ... (Your working version from previous app.py, using COMScope)
    if platform.system() != "Windows": return False, "Outlook sending only on Windows."
    try:
        with COMScope():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            if not outlook_app: return False, "Outlook app not connected."
            namespace = outlook_app.GetNamespace("MAPI"); original_item = namespace.GetItemFromID(original_entry_id)
            if not original_item or original_item.Class != 43: return False, "Original Outlook mail not found."
            reply = original_item.ReplyAll(); reply.To = to_recipients_str; reply.Subject = subject
            cleaned_body_html = body.replace('\n', '<br>'); current_html_body = reply.HTMLBody 
            if "<body>" in current_html_body.lower():
                 body_tag_end = current_html_body.lower().find("<body>") + len("<body>")
                 reply.HTMLBody = current_html_body[:body_tag_end] + f"<div>{cleaned_body_html}</div><br>" + current_html_body[body_tag_end:]
            else: reply.HTMLBody = f"<div>{cleaned_body_html}</div><br><hr>{current_html_body}"
            reply.Send()
            return True, "Outlook reply sent."
    except Exception as e: app.logger.error(f"Error sending Outlook reply: {e}", exc_info=True)
    return False, str(e)


# --- Microsoft Graph (SharePoint) Functions ---
# (Keep the get_msgraph_token, get_sharepoint_site_id, search_sharepoint_documents, 
#  get_sharepoint_document_content_text functions from the previous "SharePoint" response.
#  Ensure they use app.logger and requests - not ollama_requests for Graph calls.)
def get_msgraph_token(): # Gets token from MSAL cache for Graph API
    # This is a simplified version for local dev.
    # A full web app would have /login-microsoft and /callback-microsoft routes for interactive flow.
    ms_app = msal.PublicClientApplication(MS_GRAPH_CLIENT_ID, authority=MS_GRAPH_AUTHORITY, token_cache=ms_graph_token_cache)
    accounts = ms_app.get_accounts()
    if accounts:
        app.logger.info(f"Attempting to acquire MS Graph token silently for account: {accounts[0]['username']}")
        result = ms_app.acquire_token_silent(MS_GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            app.logger.info("MS Graph token acquired silently.")
            return result['access_token']
        else: # Silent failed, token might have expired, or needs more consent
            app.logger.warning("MS Graph silent token acquisition failed. Trying interactive (will print to console).")
            # Fallback to interactive device flow if no silent token (for dev)
            # This part is tricky to integrate directly into a Flask request smoothly without redirects
            # For a local server, you might run this part once using a separate script.
            # To try device flow here (prints code to console):
            flow = ms_app.initiate_device_flow(scopes=MS_GRAPH_SCOPES)
            if "user_code" not in flow:
                app.logger.error("Failed to create MS Graph device flow: " + flow.get("error_description", "Unknown error"))
                return None
            app.logger.info("MS Graph Device Flow initiated. Please go to: " + flow["verification_uri"] + " and enter code: " + flow["user_code"])
            # This next line will block until user authenticates or flow times out.
            try:
                result = ms_app.acquire_token_by_device_flow(flow) # This can take time
                if result and "access_token" in result:
                    app.logger.info("MS Graph token acquired via device flow.")
                    save_ms_graph_cache() # Save the new token
                    return result['access_token']
                else:
                    app.logger.error("MS Graph device flow did not return a token.")
                    return None
            except Exception as e_device_flow:
                app.logger.error(f"Error during MS Graph device flow token acquisition: {e_device_flow}")
                return None
                
    else: # No accounts in cache
        app.logger.warning("No MS Graph accounts in cache. User needs to authenticate for SharePoint features.")
        app.logger.info(f"To authenticate for MS Graph (SharePoint), you can (for dev) trigger a device flow, "
                        f"or implement /login-microsoft and /callback-microsoft routes in Flask using "
                        f"ms_app.get_authorization_request_url and ms_app.acquire_token_by_authorization_code.")
        # For now, we won't automatically trigger the device flow here to prevent server hanging on startup
        # A dedicated /login-microsoft route would be better.
    return None

# Placeholder for actual SharePoint functions, ensure they use 'requests' for HTTP calls
def get_sharepoint_site_id(access_token, site_name_to_search):
    if not access_token or not site_name_to_search: return None
    headers = {'Authorization': 'Bearer ' + access_token}
    search_url = f"https://graph.microsoft.com/v1.0/sites?search={site_name_to_search}" # Basic search
    try:
        response = requests.get(search_url, headers=headers)
        response.raise_for_status()
        sites = response.json().get('value')
        if sites:
            app.logger.info(f"Found SharePoint site '{sites[0]['name']}' with ID: {sites[0]['id']}")
            return sites[0]['id']
        app.logger.warning(f"SharePoint site '{site_name_to_search}' not found.")
    except Exception as e:
        app.logger.error(f"Error getting SharePoint site ID for '{site_name_to_search}': {e}", exc_info=True)
    return None

def search_sharepoint_documents(access_token, query_terms, site_id, drive_name="Documents", top_n=1):
    # ... (Graph API search logic as defined in previous SharePoint response, using app.logger)
    if not access_token or not site_id: return []
    headers = {'Authorization': 'Bearer ' + access_token}
    drive_id_val = None # Get drive ID first
    try:
        drive_search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$filter=name eq '{drive_name}'"
        response_drives = requests.get(drive_search_url, headers=headers)
        response_drives.raise_for_status(); drives = response_drives.json().get('value')
        if drives: drive_id_val = drives[0]['id']
        else: app.logger.warning(f"Drive '{drive_name}' not found in site {site_id}."); return []
    except Exception as e: app.logger.error(f"Error finding drive '{drive_name}': {e}"); return []

    if not drive_id_val: return []
    search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id_val}/root/search(q='{query_terms}')?$top={top_n}&$select=name,id,webUrl,file"
    app.logger.info(f"Searching SharePoint drive {drive_id_val} with query: {query_terms}")
    try:
        response = requests.get(search_url, headers=headers); response.raise_for_status()
        results = response.json().get('value', [])
        app.logger.info(f"Found {len(results)} SP docs for query '{query_terms}'.")
        return [{"name": item.get('name'), "id": item.get('id'), "site_id": site_id, "webUrl": item.get("webUrl"), "mimeType": item.get("file", {}).get("mimeType")} for item in results]
    except Exception as e: app.logger.error(f"Error searching SharePoint: {e}", exc_info=True); return []


def get_sharepoint_document_content_text(access_token, site_id, item_id, mime_type=None, item_name="UnknownFile"):
    # ... (Graph API download and parsing for TXT, DOCX, PDF as defined in previous SharePoint response, using app.logger)
    # ... This function needs python-docx and PyMuPDF (fitz)
    if not access_token or not site_id or not item_id: return None
    headers = {'Authorization': 'Bearer ' + access_token}
    download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/items/{item_id}/content" # Note: /drive/items/{item-id}
    app.logger.info(f"Downloading SP content for item: {item_id}")
    content_text = None
    try:
        response = requests.get(download_url, headers=headers, stream=True); response.raise_for_status()
        
        # Determine filename extension for parsing
        _, ext = os.path.splitext(item_name.lower())
        if not ext and mime_type: # Try to get ext from mime_type
            if mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ext = ".docx"
            elif mime_type == "application/pdf": ext = ".pdf"
            elif mime_type == "text/plain": ext = ".txt"

        if ext == ".txt" or ext == ".md":
            content_text = response.text # For text files, .text should handle encoding
        elif ext == ".docx":
            from io import BytesIO
            bytes_io = BytesIO(response.content)
            doc = DocxDocument(bytes_io)
            content_text = "\n".join([para.text for para in doc.paragraphs])
        elif ext == ".pdf":
            from io import BytesIO
            pdf_bytes = BytesIO(response.content)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            content_text = ""
            for page_num in range(len(doc)): content_text += doc.load_page(page_num).get_text()
            doc.close()
        else: app.logger.warning(f"Unsupported file type for SP content extraction: {item_name} (ext: {ext}, mime: {mime_type})")
            
        if content_text: app.logger.info(f"Extracted text (len {len(content_text)}) from SP item {item_name}.")
        return content_text
    except Exception as e: app.logger.error(f"Error getting/parsing SP doc content for item {item_id} ('{item_name}'): {e}", exc_info=True)
    return None


# --- RAG Ollama Functions (query_ollama, clean_llm_reply, get_style_examples_from_platform, draft_reply_with_rag) ---
# draft_reply_with_rag needs to be updated to include sharepoint_context_text
# (Ensure query_ollama, clean_llm_reply are present as per previous definitions)
def load_user_style_examples(csv_filepath, body_column_name, num_examples=3): # For CSV Fallback
    # ... (Same as your working version) ...
    examples = []
    try:
        script_dir = os.path.dirname(__file__)
        abs_file_path = os.path.join(script_dir, csv_filepath)
        if not os.path.exists(abs_file_path):
            app.logger.warning(f"Fallback CSV for style not found: {abs_file_path}")
            return []
        with open(abs_file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            if body_column_name not in reader.fieldnames:
                app.logger.error(f"Column '{body_column_name}' not found in CSV '{abs_file_path}'. Fields: {reader.fieldnames}")
                return []
            all_bodies = [row[body_column_name] for row in reader if row[body_column_name] and len(row[body_column_name]) > 30]
            if not all_bodies: return []
            examples = random.sample(all_bodies, min(len(all_bodies), num_examples))
            app.logger.info(f"Loaded {len(examples)} style examples from CSV '{abs_file_path}'.")
    except Exception as e:
        app.logger.error(f"Error loading style examples from CSV '{csv_filepath}': {e}", exc_info=True)
    return [ex_body for ex_body in examples if isinstance(ex_body, str) or (isinstance(ex_body, dict) and ex_body.get(body_column_name))]


def get_style_examples_from_platform(platform_type, count=NUM_STYLE_EXAMPLES):
    # ... (Same as before, ensure it uses your working internal fetchers) ...
    style_example_bodies = []
    app.logger.info(f"Attempting to fetch {count} sent items for style from {platform_type}")
    if platform_type == "gmail":
        gmail_service = get_gmail_service() 
        if gmail_service:
            fetched_items = fetch_gmail_emails_internal(gmail_service, 'SENT', count, for_style=True)
            style_example_bodies = [item['body'] for item in fetched_items if item.get('body')]
    elif platform_type == "outlook":
        fetched_items = fetch_outlook_emails_internal(5, count, for_style=True) 
        style_example_bodies = [item['body'] for item in fetched_items if item.get('body')]
    
    if not style_example_bodies: 
        app.logger.warning(f"Live fetch for {platform_type} sent items failed or returned empty. Trying CSV fallback.")
        csv_to_try = USER_SENT_GMAIL_CSV if platform_type == "gmail" else USER_SENT_OUTLOOK_CSV
        loaded_from_csv = load_user_style_examples(csv_to_try, CSV_BODY_COLUMN_NAME, count)
        style_example_bodies = [ex if isinstance(ex, str) else ex.get(CSV_BODY_COLUMN_NAME, '') for ex in loaded_from_csv]
        style_example_bodies = [s for s in style_example_bodies if s] 

    app.logger.info(f"Loaded {len(style_example_bodies)} style examples for RAG on platform {platform_type}.")
    return style_example_bodies

def draft_reply_with_rag(user_name_for_prompt, 
                         incoming_email_platform, 
                         incoming_email_sender, 
                         incoming_email_subject, 
                         incoming_email_body, 
                         style_examples_list, 
                         sharepoint_context_text=""): # New parameter for SP context
    # ... (Prompt construction now includes sharepoint_prompt_addition as shown in the previous response) ...
    style_instruction_block = ""
    if not style_examples_list:
        app.logger.info(f"No style examples for RAG prompt (platform: {incoming_email_platform}).")
        style_instruction_block = "Please draft a professional and helpful reply."
    else:
        style_examples_text = ""
        for i, example in enumerate(style_examples_list):
            style_examples_text += f"Example {i+1} of {user_name_for_prompt}'s writing (from {incoming_email_platform}):\n{example}\n---\n"
        style_instruction_block = f"""The following email excerpts are provided ONLY to demonstrate {user_name_for_prompt}'s typical writing style...
                               **DO NOT use the topics... from these style examples...**
                               Your reply's content must be based **SOLELY** on the new incoming email...
                               --- Start of Writing Style Examples ---
                               {style_examples_text}--- End of Writing Style Examples ---
                               When drafting the reply... emulate the writing style of {user_name_for_prompt}...
                               ...substance and facts... from the new incoming email's content..."""

    sharepoint_prompt_addition = ""
    if sharepoint_context_text:
        app.logger.info(f"Adding SharePoint context to RAG prompt (len: {len(sharepoint_context_text)}).")
        sharepoint_prompt_addition = f"""\n--- Relevant Information from Company Documents (SharePoint) ---
{sharepoint_context_text[:3000]} 
--- End of Company Document Information ---

When drafting your reply, please consider and utilize the relevant information from the company documents provided above to make your response more accurate and informed, if applicable to the incoming email's query.
"""
    prompt = f"""You are an AI assistant helping {user_name_for_prompt} draft a reply to an important email.
               **Your Primary Task:** Reply to the **new incoming email**...
               **New Incoming Email Details:** Platform: {incoming_email_platform}, From: "{incoming_email_sender}", Subject: "{incoming_email_subject}", Body:\n{incoming_email_body}\n---
               {sharepoint_prompt_addition}
               **Writing Style Guidance:**\n{style_instruction_block}
               **Instructions for the Reply Draft:**
               1. Address all points... using information from company documents if relevant...
               2. Focus on reply body...
               3. Salutation/Sign-off: Do not add...
               4. Accuracy...
               5. Style Adherence...
               Draft the reply body for the **new incoming email** now:"""
    return query_ollama(prompt)

# --- Main API Endpoints ---
@app.route('/auth-status', methods=['GET'])
def api_auth_status():
    # ... (Same as before, ensure it correctly calls get_gmail_service and tests Outlook dispatch) ...
    gmail_ok = False; outlook_ok = False
    if get_gmail_service(interactive_auth_if_needed=False): gmail_ok = True
    if platform.system() == "Windows":
        try:
            with COMScope():
                if win32com.client.Dispatch("Outlook.Application"): outlook_ok = True
        except: pass
    return jsonify({"gmail": gmail_ok, "outlook": outlook_ok, "sharepoint_ready": bool(get_msgraph_token() is not None)})


@app.route('/initiate-gmail-auth', methods=['GET'])
def initiate_gmail_auth_route():
    # ... (Same as before) ...
    if get_gmail_service(interactive_auth_if_needed=True):
        return jsonify({"status": "success", "message": "Gmail auth flow initiated."})
    return jsonify({"status": "error", "message": "Gmail auth failed."}), 500

@app.route('/initiate-microsoft-auth', methods=['GET']) # New for Graph/SharePoint
def initiate_microsoft_auth():
    ms_app = msal.PublicClientApplication(MS_GRAPH_CLIENT_ID, authority=MS_GRAPH_AUTHORITY, token_cache=ms_graph_token_cache)
    flow = ms_app.initiate_device_flow(scopes=MS_GRAPH_SCOPES)
    if "user_code" not in flow:
        app.logger.error("Failed to create MS Graph device flow: " + flow.get("error_description", "Unknown error"))
        return jsonify({"error": "Could not initiate Microsoft auth device flow."}), 500
    
    app.logger.info("MS Graph Device Flow initiated. User should go to: " + flow["verification_uri"] + " and enter code: " + flow["user_code"])
    # Store flow in session to check against later if needed, or just let user complete it.
    # For this simple case, we assume user completes it and then /auth-status will reflect it.
    return jsonify({
        "message": "Microsoft authentication (for SharePoint/Graph) initiated. Please follow instructions below.",
        "verification_uri": flow["verification_uri"],
        "user_code": flow["user_code"],
        "expires_in": flow["expires_in"]
    })


@app.route('/emails', methods=['GET'])
def get_emails_route():
    # ... (Same as before, ensure robust date parsing for sorting) ...
    all_emails = []
    gmail_service = get_gmail_service() 
    if gmail_service: all_emails.extend(fetch_gmail_emails_internal(gmail_service, 'INBOX', 15))
    all_emails.extend(fetch_outlook_emails_internal(6, 15))
    try:
        def parse_date_robust(date_str_or_obj):
            # ... (same robust date parser)
            if isinstance(date_str_or_obj, datetime): return date_str_or_obj.replace(tzinfo=None)
            if not date_str_or_obj: return datetime.min
            date_str = str(date_str_or_obj) 
            try: return datetime.strptime(date_str.split(' (')[0].strip(), "%a, %d %b %Y %H:%M:%S %z").replace(tzinfo=None)
            except ValueError: pass
            try: return datetime.strptime(date_str.split(' (')[0].strip(), "%d %b %Y %H:%M:%S %z").replace(tzinfo=None) 
            except ValueError: pass
            try: return datetime.fromisoformat(date_str.replace('Z', '+00:00')).replace(tzinfo=None)
            except ValueError: pass
            app.logger.warning(f"Unparseable date string: {date_str}")
            return datetime.min 
        all_emails.sort(key=lambda x: parse_date_robust(x.get('date')), reverse=True)
    except Exception as e: app.logger.warning(f"Could not sort emails by date: {e}.")
    return jsonify(all_emails)

@app.route('/email-details', methods=['GET'])
def get_single_email_details_route():
    # ... (Same as before, using the updated internal detail fetchers) ...
    platform_type = request.args.get('platform'); email_id = request.args.get('id')
    if not platform_type or not email_id: return jsonify({"error": "Missing platform or email ID"}), 400
    details = None
    if platform_type == 'gmail':
        gmail_service = get_gmail_service()
        if gmail_service: details = get_gmail_email_details_internal(gmail_service, email_id)
        else: return jsonify({"error": "Gmail service N/A."}), 503
    elif platform_type == 'outlook':
        details = get_outlook_email_details_internal(email_id) 
    if details: return jsonify(details)
    return jsonify({"error": f"Could not fetch details for {platform_type} ID {email_id}. Check server logs."}), 404


@app.route('/draft-ai-reply', methods=['POST'])
def draft_ai_reply_endpoint_route():
    data = request.get_json()
    platform_type = data.get('platform')
    original_sender = data.get('sender')
    original_subject = data.get('subject')
    original_body = data.get('body')
    user_name_for_prompt = data.get('userName', USER_NAME)

    if not all([platform_type, original_subject is not None, original_body is not None]):
        return jsonify({"error": "Missing platform, subject, or body for drafting reply"}), 400

    app.logger.info(f"Drafting reply for {platform_type} email. Subject: {original_subject[:50]}")
    
    style_examples = get_style_examples_from_platform(platform_type, NUM_STYLE_EXAMPLES)
    
    # --- SharePoint Document Retrieval ---
    sharepoint_text_context = ""
    sharepoint_docs_found_names = [] # To inform UI
    # Decide if SP search should happen for gmail, outlook, or both
    # For now, let's assume it's relevant for any type of email if MS Graph is authed
    ms_graph_token = get_msgraph_token()
    if ms_graph_token:
        app.logger.info("Microsoft Graph token available, attempting SharePoint search.")
        target_site_id = get_sharepoint_site_id(ms_graph_token, SHAREPOINT_SITE_NAME)
        if target_site_id:
            search_terms_from_email = f"{original_subject} {original_body[:500]}" # Simple keyword source
            relevant_docs_info = search_sharepoint_documents(ms_graph_token, search_terms_from_email, target_site_id, SHAREPOINT_DEFAULT_DRIVE_NAME, top_n=1)
            
            if relevant_docs_info:
                # For now, take content from the first relevant doc
                doc_info = relevant_docs_info[0]
                sharepoint_docs_found_names.append(doc_info.get('name'))
                app.logger.info(f"Attempting to extract content from SP doc: {doc_info.get('name')}")
                doc_content = get_sharepoint_document_content_text(ms_graph_token, doc_info['site_id'], doc_info['id'], doc_info.get('mimeType'), doc_info.get('name'))
                if doc_content:
                    sharepoint_text_context = doc_content[:3000] # Limit context size
                    app.logger.info(f"Extracted content (len {len(sharepoint_text_context)}) from SP doc {doc_info.get('name')}.")
            else:
                app.logger.info(f"No relevant SP documents found for query: '{search_terms_from_email[:50]}...'")
        else:
            app.logger.warning(f"Could not get SP Site ID for '{SHAREPOINT_SITE_NAME}'. Skipping SP search.")
    else:
        app.logger.info("MS Graph token not available, skipping SharePoint document search.")
    # --- End SharePoint ---
    
    raw_draft = draft_reply_with_rag(
        user_name_for_prompt, platform_type, original_sender, 
        original_subject, original_body, style_examples,
        sharepoint_text_context # Pass new context
    )
    
    if raw_draft:
        cleaned_draft = clean_llm_reply(raw_draft)
        return jsonify({"draft": cleaned_draft, "sharepoint_docs_found": sharepoint_docs_found_names})
    else:
        return jsonify({"error": "Failed to generate draft from AI service"}), 500

@app.route('/send-platform-reply', methods=['POST'])
def send_platform_reply_endpoint_route():
    # ... (Same as before, ensure it calls your working internal send functions) ...
    data = request.get_json(); platform_type = data.get('platform'); original_message_id = data.get('originalMessageId'); original_thread_id = data.get('originalThreadId'); to_recipients = data.get('to'); subject = data.get('subject'); body = data.get('body'); in_reply_to_header = data.get('inReplyToHeader'); references_header = data.get('referencesHeader')
    if not all([platform_type, original_message_id, to_recipients, subject, body is not None]): return jsonify({"error": "Missing fields for sending"}), 400
    success = False; message_or_status = ""
    if platform_type == 'gmail':
        gmail_service = get_gmail_service()
        if gmail_service: success, message_or_status = send_gmail_reply_internal(gmail_service, to_recipients, subject, body, original_thread_id, in_reply_to_header, references_header)
        else: message_or_status = "Gmail service N/A."
    elif platform_type == 'outlook':
        success, message_or_status = send_outlook_reply_internal(original_message_id, to_recipients, subject, body)
    else: message_or_status = f"Platform '{platform_type}' N/A."
    if success: return jsonify({"status": "success", "message": message_or_status})
    return jsonify({"status": "error", "message": message_or_status}), 500

if __name__ == '__main__':
    print("Starting Flask server: Unified Email RAG Drafter with SharePoint...")
    # ... (Startup messages from before) ...
    if not MS_GRAPH_CLIENT_ID or not MS_GRAPH_AUTHORITY:
        print("WARNING: MS_GRAPH_CLIENT_ID or MS_GRAPH_AUTHORITY not set in .env file. SharePoint features will require manual auth or will fail.")
    else:
        print(f"MS Graph Client ID: {MS_GRAPH_CLIENT_ID[:5]}...")
        print(f"SharePoint Site Target: {SHAREPOINT_SITE_NAME}")

    print(f"Backend accessible at http://localhost:5000")
    use_threading = False if platform.system() == "Windows" else True
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=use_threading, use_reloader=False)