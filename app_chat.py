# app_chat.py (Updated for better MSAL token debugging)
from flask import Flask, request, jsonify, session
from flask_cors import CORS
import os
import platform
import json
import requests # For MS Graph and Ollama
import re
from dotenv import load_dotenv
import numpy as np

# MSAL for Microsoft Graph (SharePoint)
import msal
import atexit

# Document parsing
from docx import Document as DocxDocument
import fitz  # PyMuPDF for PDF text extraction

# Embeddings and Vector DB
from sentence_transformers import SentenceTransformer
import faiss

# --- Load Environment Variables ---
load_dotenv()

# --- Flask App Setup ---
app = Flask(__name__)
CORS(app, supports_credentials=True)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(32))

# --- Configuration ---
OLLAMA_API_URL = os.getenv("OLLAMA_API_URL", "http://localhost:11434/api/generate")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "deepseek-r1:14b")
USER_NAME = os.getenv("USER_NAME", "Khurram")
SCRIPT_DIR_APP = os.path.dirname(os.path.abspath(__file__))
# Microsoft Graph (SharePoint) Config
MS_GRAPH_CLIENT_ID = os.getenv('MS_GRAPH_CLIENT_ID')
MS_GRAPH_AUTHORITY = os.getenv('MS_GRAPH_AUTHORITY')
MS_GRAPH_SCOPES = ["User.Read", "Sites.Read.All", "Files.Read.All"] # Ensure 'offline_access' is NOT here
MS_GRAPH_TOKEN_CACHE_FILE = os.path.join(SCRIPT_DIR_APP, "token_cache_ms_graph_chat.bin")

SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_LIBRARY_NAME = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documents")
SHAREPOINT_FOLDER_PATH = os.getenv("SHAREPOINT_FOLDER_PATH", "")

FAISS_INDEX_PATH = "sharepoint_faiss.index"
FAISS_METADATA_PATH = "sharepoint_metadata.json"
EMBEDDING_MODEL_NAME = 'all-MiniLM-L6-v2'
embedding_model = None
faiss_index = None
doc_metadata = []

# --- MSAL Token Cache for Graph API ---
ms_graph_token_cache = msal.SerializableTokenCache()
if os.path.exists(MS_GRAPH_TOKEN_CACHE_FILE):
    try:
        app.logger.info(f"Attempting to load MS Graph token cache from {MS_GRAPH_TOKEN_CACHE_FILE}...")
        cache_content = open(MS_GRAPH_TOKEN_CACHE_FILE, "r").read()
        if cache_content:
            ms_graph_token_cache.deserialize(cache_content)
            app.logger.info(f"MS Graph token cache deserialized. Cache size: {len(cache_content)}")
            
            # Immediately check accounts after deserializing
            temp_ms_app_on_load = msal.PublicClientApplication(MS_GRAPH_CLIENT_ID, authority=MS_GRAPH_AUTHORITY, token_cache=ms_graph_token_cache)
            cached_accounts_on_load = temp_ms_app_on_load.get_accounts()
            if cached_accounts_on_load:
                app.logger.info(f"Accounts found in loaded cache at startup: {[acc.get('username') for acc in cached_accounts_on_load]}")
            else:
                app.logger.warning("No accounts found in loaded MS Graph token cache at startup. Cache might be empty or not recognized.")
        else:
            app.logger.warning(f"MS Graph token cache file '{MS_GRAPH_TOKEN_CACHE_FILE}' is empty.")
    except Exception as e:
        app.logger.error(f"Error loading MS Graph token cache from {MS_GRAPH_TOKEN_CACHE_FILE}: {e}. A new cache may be created if auth is triggered.", exc_info=True)
else:
    app.logger.warning(f"MS Graph token cache file '{MS_GRAPH_TOKEN_CACHE_FILE}' not found. User needs to authenticate.")

def save_ms_graph_cache():
    if ms_graph_token_cache.has_state_changed:
        with open(MS_GRAPH_TOKEN_CACHE_FILE, "w") as cache_file:
            cache_file.write(ms_graph_token_cache.serialize())
        app.logger.info(f"MS Graph token cache state changed and saved to {MS_GRAPH_TOKEN_CACHE_FILE}")
    else:
        app.logger.info(f"MS Graph token cache state has not changed. Not saving.")
atexit.register(save_ms_graph_cache)

def get_ms_graph_token_for_chat():
    if not MS_GRAPH_CLIENT_ID or not MS_GRAPH_AUTHORITY:
        app.logger.error("MS_GRAPH_CLIENT_ID or MS_GRAPH_AUTHORITY is not configured in .env file.")
        return None

    ms_app = msal.PublicClientApplication(
        MS_GRAPH_CLIENT_ID, authority=MS_GRAPH_AUTHORITY, token_cache=ms_graph_token_cache
    )
    accounts = ms_app.get_accounts()

    if accounts:
        app.logger.info(f"Found {len(accounts)} MS Graph account(s) in current cache instance. Usernames: {[acc.get('username') for acc in accounts]}")
        app.logger.info(f"Attempting to acquire MS Graph token silently for account: {accounts[0].get('username')} with scopes: {MS_GRAPH_SCOPES}")
        
        result = ms_app.acquire_token_silent(MS_GRAPH_SCOPES, account=accounts[0])
        
        if result and "access_token" in result:
            app.logger.info("MS Graph token acquired silently.")
            # Log token expiry if available: result.get("expires_in") or result.get("ext_expires_in")
            if "expires_in" in result:
                 app.logger.info(f"Access token expires in: {result['expires_in']} seconds")
            if "ext_expires_in" in result: # Extended expiry for Graph API
                 app.logger.info(f"Extended access token expires in: {result['ext_expires_in']} seconds")

            save_ms_graph_cache() # Save cache in case refresh token was used and cache state changed
            return result['access_token']
        else: 
            app.logger.warning("MS Graph silent token acquisition failed.")
            if result:
                app.logger.error(f"MS Graph acquire_token_silent specific error: {result.get('error')}, description: {result.get('error_description')}")
                app.logger.debug(f"Full silent acquisition result (if failed): {result}")
            else:
                app.logger.error("MS Graph acquire_token_silent returned None. This usually means no matching token was found or it was expired and no refresh token was available/usable.")
            app.logger.error("The application will not attempt interactive auth here. Please ensure 'generate_token_graph.py' was run successfully and recently.")
    else:
        app.logger.warning("No MS Graph accounts found in MSAL token cache. User needs to authenticate using generate_token_graph.py.")

    # This final log line is the one you're seeing in your console
    app.logger.error("MS Graph token not available or silent acquisition failed. Please run authentication utility (e.g., generate_token_graph.py from previous steps).")
    return None


# --- SharePoint Document Processing Functions (get_site_id, get_drive_id, list_files_in_sharepoint_folder_recursive, get_sp_doc_content, chunk_text) ---
# (These functions remain the same as the previous "SharePoint" response)
# For brevity, assuming they are correctly defined as before and use app.logger.
MS_GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0'
def get_site_id(access_token, site_name_to_search):
    if not access_token or not site_name_to_search: 
        return None
    
    headers = {'Authorization': 'Bearer ' + access_token}
    search_url = f"{MS_GRAPH_API_BASE}/sites?search={site_name_to_search}" 
    
    try:
        response = requests.get(search_url, headers=headers, timeout=10)
        # Log the full response for debugging
        app.logger.info(f"SharePoint site search response: Status={response.status_code}, Body={response.text}")
        
        response.raise_for_status()
        sites = response.json().get('value')
        if sites:
            app.logger.info(f"Found SharePoint site '{sites[0]['name']}' with ID: {sites[0]['id']}")
            return sites[0]['id']
        app.logger.warning(f"SharePoint site '{site_name_to_search}' not found.")
    except Exception as e:
        app.logger.error(f"Error getting SharePoint site ID for '{site_name_to_search}': {e}", exc_info=True)
    return None

def get_drive_id(access_token, site_id, drive_name):
    # ... (same as before)
    if not access_token or not site_id or not drive_name: return None
    headers = {'Authorization': 'Bearer ' + access_token}
    url = f"{MS_GRAPH_API_BASE}/sites/{site_id}/drives?$filter=name eq '{drive_name}'"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        drives = response.json().get('value')
        if drives:
            app.logger.info(f"Found Drive ID: {drives[0]['id']} for drive '{drive_name}' in site {site_id}")
            return drives[0]['id']
        app.logger.warning(f"Drive '{drive_name}' not found in site {site_id}.")
    except Exception as e:
        app.logger.error(f"Error finding drive '{drive_name}' in site {site_id}: {e}", exc_info=True)
    return None

def list_files_in_sharepoint_folder_recursive(access_token, site_id, drive_id, item_id="root", current_path=""):
    # ... (same as before)
    if not access_token or not site_id or not drive_id: return []
    headers = {'Authorization': 'Bearer ' + access_token}
    # Select fewer properties to speed up if only name, id, file, folder, webUrl are needed
    url = f"{MS_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}/children?$select=name,id,file,folder,webUrl"
    files_list = []
    page_count = 0
    while url:
        page_count +=1
        app.logger.debug(f"Fetching SP children from: {url.split('?')[0]}, page: {page_count}")
        try:
            response = requests.get(url, headers=headers, timeout=15); response.raise_for_status()
            items_page = response.json()
            items = items_page.get('value', [])
            for item in items:
                item_name = item.get('name'); full_path = os.path.join(current_path, item_name) if current_path else item_name
                if item.get('file') and item_name:
                    file_type = item.get('file', {}).get('mimeType', '').lower()
                    if item_name.lower().endswith(('.docx','.pdf','.txt')) or 'officedocument.wordprocessingml' in file_type or 'application/pdf' in file_type or 'text/plain' in file_type:
                        files_list.append({"name": item_name, "id": item.get('id'), "path": full_path, "webUrl": item.get('webUrl'), "mimeType": file_type})
                elif item.get('folder'):
                     app.logger.debug(f"Descending into SP folder: {full_path}")
                     files_list.extend(list_files_in_sharepoint_folder_recursive(access_token, site_id, drive_id, item.get('id'), full_path))
            url = items_page.get('@odata.nextLink') # For pagination
        except Exception as e:
            app.logger.error(f"Error listing files in SP folder (item {item_id}, path {current_path}, page {page_count}): {e}", exc_info=True); break 
    return files_list


def get_sp_doc_content(access_token, site_id, item_id, item_name, mime_type):
    # ... (same as before, using python-docx and PyMuPDF)
    if not access_token or not site_id or not item_id: return None
    headers = {'Authorization': 'Bearer ' + access_token}
    # Using /beta/drive/items/{item-id}/preview with POST can sometimes give text, but /content is more direct
    download_url = f"{MS_GRAPH_API_BASE}/sites/{site_id}/drives/items/{item_id}/content"
    app.logger.info(f"Downloading SP content for item: {item_name} (ID: {item_id})")
    content_text = None
    try:
        response = requests.get(download_url, headers=headers, stream=True, timeout=30); response.raise_for_status()
        file_ext = os.path.splitext(item_name.lower())[1]
        if not ext and mime_type: 
            if mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ext = ".docx"
            elif mime_type == "application/pdf": ext = ".pdf"
            elif mime_type == "text/plain": ext = ".txt"

        if ext == ".txt" or mime_type == "text/plain":
            content_text = response.text 
        elif ext == ".docx" or "officedocument.wordprocessingml" in mime_type:
            from io import BytesIO
            doc = DocxDocument(BytesIO(response.content))
            content_text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        elif ext == ".pdf" or mime_type == "application/pdf":
            from io import BytesIO
            with BytesIO(response.content) as pdf_stream:
                doc = fitz.open(stream=pdf_stream, filetype="pdf")
                content_text = "".join([page.get_text() + "\n" for page in doc])
                doc.close()
        else: app.logger.warning(f"Unsupported file type for SP content extraction: {item_name} (ext: {ext}, mime: {mime_type})")
        if content_text: app.logger.info(f"Extracted text (len {len(content_text)}) from SP item {item_name}.")
        return " ".join(content_text.split()) if content_text else None # Normalize whitespace
    except Exception as e: app.logger.error(f"Error getting/parsing SP doc content for item {item_id} ('{item_name}'): {e}", exc_info=True)
    return None

def chunk_text(text, chunk_size=1000, chunk_overlap=100):
    # ... (same as before)
    if not text: return []
    # A more robust chunker would split by sentences/paragraphs first.
    # For now, simple sliding window.
    return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size - chunk_overlap)]


# --- FAISS Indexing and RAG Functions (load_embedding_model, build_or_load_faiss_index, query_ollama, clean_llm_reply, draft_reply_with_rag) ---
# (These functions are mostly the same from the previous SharePoint response)
# Ensure load_embedding_model is called before FAISS operations if model is not yet loaded.

def load_embedding_model():
    global embedding_model
    if embedding_model is None:
        try:
            app.logger.info(f"Loading sentence transformer model: {EMBEDDING_MODEL_NAME}...")
            embedding_model = SentenceTransformer(EMBEDDING_MODEL_NAME)
            app.logger.info("Embedding model loaded.")
        except Exception as e:
            app.logger.error(f"Failed to load embedding model: {e}", exc_info=True)
            # raise # Or handle gracefully
    return embedding_model

def build_or_load_faiss_index(force_rebuild=False):
    global faiss_index, doc_metadata
    script_dir = os.path.dirname(__file__)
    faiss_path = os.path.join(script_dir, FAISS_INDEX_PATH)
    meta_path = os.path.join(script_dir, FAISS_METADATA_PATH)

    if os.path.exists(faiss_path) and os.path.exists(meta_path) and not force_rebuild:
        try:
            app.logger.info(f"Loading existing FAISS index from {faiss_path}")
            faiss_index = faiss.read_index(faiss_path)
            with open(meta_path, 'r', encoding='utf-8') as f:
                doc_metadata = json.load(f)
            app.logger.info(f"Loaded FAISS index with {faiss_index.ntotal} vectors and {len(doc_metadata)} metadata entries.")
            load_embedding_model() 
            return True, "Knowledgebase loaded from disk."
        except Exception as e:
            app.logger.error(f"Error loading FAISS index/metadata: {e}. Will try to rebuild.", exc_info=True)
            faiss_index = None; doc_metadata = []
    return False, "Knowledgebase not loaded. Please update."

@app.route('/update-knowledgebase', methods=['POST'])
def update_knowledgebase():
    global faiss_index, doc_metadata # To modify global instances
    app.logger.info("Update Knowledgebase request received.")
    
    access_token = get_ms_graph_token_for_chat() # This is where the error you saw originates
    if not access_token:
        app.logger.error("MS Graph authentication failed. Cannot access SharePoint for knowledge base update.")
        return jsonify({"error": "Microsoft Graph not authenticated. Cannot access SharePoint. Please check server logs and ensure token cache is valid or run auth utility."}), 401

    # Now, if access_token is successfully retrieved, proceed:
    app.logger.info("MS Graph token acquired, proceeding with SharePoint document fetch.")
    site_id = get_site_id(access_token, SHAREPOINT_SITE_NAME)
    if not site_id:
        return jsonify({"error": f"Could not find SharePoint site: {SHAREPOINT_SITE_NAME}"}), 404
    
    drive_id = get_drive_id(access_token, site_id, SHAREPOINT_LIBRARY_NAME)
    if not drive_id:
         return jsonify({"error": f"Could not find SharePoint library: {SHAREPOINT_LIBRARY_NAME}"}), 404

    item_id_to_list = "root"
    if SHAREPOINT_FOLDER_PATH:
        folder_url = f"{MS_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root:/{SHAREPOINT_FOLDER_PATH}"
        headers = {'Authorization': 'Bearer ' + access_token}
        try:
            folder_item_resp = requests.get(folder_url, headers=headers, timeout=10); folder_item_resp.raise_for_status()
            item_id_to_list = folder_item_resp.json().get('id', 'root')
        except Exception as e_folder:
            app.logger.error(f"Could not resolve SP folder path '{SHAREPOINT_FOLDER_PATH}': {e_folder}. Indexing from library root.")
            # Fall through to item_id_to_list = "root"

    sharepoint_files = list_files_in_sharepoint_folder_recursive(access_token, site_id, drive_id, item_id_to_list)
    if not sharepoint_files:
        return jsonify({"message": "No compatible documents found in SharePoint or error fetching.", "indexed_count": 0}), 200

    all_chunks = []; new_doc_metadata = []
    model = load_embedding_model()
    if not model: return jsonify({"error": "Embedding model failed to load."}), 500

    for i, file_info in enumerate(sharepoint_files):
        app.logger.info(f"Processing document {i+1}/{len(sharepoint_files)}: {file_info['name']}")
        content = get_sp_doc_content(access_token, site_id, file_info['id'], file_info['name'], file_info.get('mimeType'))
        if content:
            chunks = chunk_text(content)
            for chunk_idx, chunk_text_content in enumerate(chunks): # Renamed to avoid conflict
                all_chunks.append(chunk_text_content)
                new_doc_metadata.append({
                    "source_doc_name": file_info['name'],
                    "source_doc_path": file_info.get('path', file_info['name']),
                    "webUrl": file_info.get('webUrl'),
                    "chunk_text": chunk_text_content, # STORE FULL CHUNK HERE
                    "chunk_id": f"{file_info['id']}_{chunk_idx}" 
                })
    if not all_chunks: return jsonify({"message": "No text content extracted.", "indexed_count": 0}), 200

    app.logger.info(f"Generating embeddings for {len(all_chunks)} text chunks...")
    embeddings = model.encode(all_chunks, show_progress_bar=True)
    dimension = embeddings.shape[1]
    new_faiss_index = faiss.IndexFlatL2(dimension); new_faiss_index.add(np.array(embeddings).astype('float32'))
    faiss_index = new_faiss_index; doc_metadata = new_doc_metadata # Update globals

    script_dir = os.path.dirname(__file__) # Ensure paths are relative to app.py
    faiss_path = os.path.join(script_dir, FAISS_INDEX_PATH)
    meta_path = os.path.join(script_dir, FAISS_METADATA_PATH)
    try:
        faiss.write_index(faiss_index, faiss_path)
        with open(meta_path, 'w', encoding='utf-8') as f: json.dump(doc_metadata, f, indent=4)
        return jsonify({"message": f"Knowledgebase updated. Indexed {faiss_index.ntotal} chunks.", "indexed_chunk_count": faiss_index.ntotal})
    except Exception as e:
        return jsonify({"error": f"Error saving knowledgebase: {str(e)}"}), 500


@app.route('/chat-with-sp-docs', methods=['POST'])
def chat_with_sp_docs():
    global faiss_index, doc_metadata, embedding_model
    if faiss_index is None or not doc_metadata:
        loaded_ok, msg = build_or_load_faiss_index()
        if not loaded_ok: return jsonify({"error": msg, "response": "", "sources": []}), 400
    if embedding_model is None: return jsonify({"error": "Embedding model not loaded.", "response": "", "sources": []}), 500

    data = request.get_json(); user_query = data.get('query'); chat_history_str = data.get('history', "")
    if not user_query: return jsonify({"error": "Query missing."}), 400
    
    app.logger.info(f"Chat query: {user_query}")
    query_embedding = embedding_model.encode([user_query])[0]
    K = 3; distances, indices = faiss_index.search(np.array([query_embedding]).astype('float32'), K)
    
    retrieved_chunks_texts = []; retrieved_sources = []
    for i, idx in enumerate(indices[0]):
        if idx != -1:
            chunk_meta = doc_metadata[idx]
            retrieved_chunks_texts.append(chunk_meta.get("chunk_text", "")) # Use full chunk text
            source_info = {"name": chunk_meta.get("source_doc_name"), "path": chunk_meta.get("source_doc_path"), "webUrl": chunk_meta.get("webUrl")}
            if source_info not in retrieved_sources : retrieved_sources.append(source_info)
    
    rag_context = "\n\n---\n\n".join(retrieved_chunks_texts)
    history_prefix = f"Previous conversation:\n{chat_history_str}\n\n" if chat_history_str else ""
    prompt = f"""{history_prefix}User query: {user_query}
               You are an AI assistant answering based on SharePoint documents.
               Provided context:
               --- Context Start ---\n{rag_context if rag_context else "No specific context found."}\n--- Context End ---
               Based ONLY on provided context, user query, and chat history, answer. If context lacks answer, say so.
               Your answer:"""
    raw_llm_response = query_ollama(prompt)
    cleaned_response = clean_llm_reply(raw_llm_response if raw_llm_response else "Sorry, I could not generate a response.")
    return jsonify({"response": cleaned_response, "sources": retrieved_sources})

@app.route('/list-indexed-documents', methods=['GET'])
def list_indexed_documents():
    # ... (Same as before, loads FAISS index if not already) ...
    global doc_metadata
    if not doc_metadata:
        loaded_ok, _ = build_or_load_faiss_index()
        if not loaded_ok:
            return jsonify({"error": "Knowledgebase not loaded. Please update.", "documents": []}), 400
    unique_docs = {}
    for meta in doc_metadata:
        doc_name = meta.get("source_doc_name")
        if doc_name and doc_name not in unique_docs:
            unique_docs[doc_name] = {"name": doc_name, "path": meta.get("source_doc_path"), "webUrl": meta.get("webUrl")}
    return jsonify({"documents": list(unique_docs.values())})

# --- Ollama/RAG Helper functions (query_ollama, clean_llm_reply - ensure they are defined as provided before) ---
def query_ollama(prompt, model_name=OLLAMA_MODEL):
    # ... (Same robust version from previous replies) ...
    try:
        payload = {"model": model_name, "prompt": prompt, "stream": False, "options": {"temperature": 0.5}}
        app.logger.info(f"Querying Ollama model: {model_name}. Prompt length: {len(prompt)}")
        app.logger.debug(f"Ollama Prompt (first 300 chars): {prompt[:300]}...")
        response = ollama_requests.post(OLLAMA_API_URL, json=payload, timeout=180)
        response.raise_for_status()
        response_data = response.json()
        app.logger.info("Ollama response received.")
        if "response" in response_data:
            return response_data["response"].strip()
        elif "error" in response_data:
            app.logger.error(f"Ollama API error: {response_data['error']}")
            return f"[Ollama Error: {response_data['error']}]"
        app.logger.error(f"Ollama response unexpected: {response_data}")
        return "[Ollama Error: Unexpected response format]"
    except ollama_requests.exceptions.ConnectionError:
        app.logger.error(f"Could not connect to Ollama API at {OLLAMA_API_URL.rsplit('/api/', 1)[0]}. Ensure Ollama is running.")
    except ollama_requests.exceptions.Timeout:
        app.logger.error("Request to Ollama API timed out.")
    except Exception as e:
        app.logger.error(f"Unexpected error querying Ollama: {e}", exc_info=True)
    return None

def clean_llm_reply(raw_reply):
    # ... (Same robust version from previous replies) ...
    if not raw_reply: return ""
    cleaned = raw_reply
    think_block_pattern = r"<think>.*?</think>\s*"
    cleaned = re.sub(think_block_pattern, "", cleaned, flags=re.DOTALL).strip()
    common_intros = [
        "Certainly! Here's a draft of the reply based on the tone and style demonstrated in the examples:",
        "Certainly! Here's a draft of the reply based on the tone and style demonstrated in the provided examples:",
        "Certainly! Here's a draft of the reply:", "Certainly! Here's a draft:",
        "Okay, here's a draft:", "Here's a draft based on your input:",
        "Here is a draft reply:", "Draft reply:", "Reply:"
    ] 
    for intro in common_intros:
        if cleaned.lower().startswith(intro.lower()):
            cleaned = cleaned[len(intro):].strip()
            if cleaned.startswith("---"): 
                parts = cleaned.split("---", 1)
                if len(parts) > 1: cleaned = parts[1].strip()
            break
    common_trailers_regex = r"\s*---\s*(This draft maintains|This reply attempts to|This response aims to|I hope this draft is helpful).*"
    cleaned = re.sub(common_trailers_regex, "", cleaned, flags=re.DOTALL | re.IGNORECASE).strip()
    return cleaned


# --- Initial Load ---
build_or_load_faiss_index() # Attempt to load FAISS index on startup

if __name__ == '__main__':
    if not all([MS_GRAPH_CLIENT_ID, MS_GRAPH_AUTHORITY, SHAREPOINT_SITE_NAME]):
        app.logger.critical("CRITICAL: MS_GRAPH_CLIENT_ID, MS_GRAPH_AUTHORITY, or SHAREPOINT_SITE_NAME not set in .env. SharePoint features WILL FAIL.")
        print("CRITICAL: MS_GRAPH_CLIENT_ID, MS_GRAPH_AUTHORITY, or SHAREPOINT_SITE_NAME not set in .env. SharePoint features WILL FAIL.")
    
    print("Starting Flask server for SharePoint RAG Chat...")
    print(f"Ollama Model: {OLLAMA_MODEL}")
    print(f"SharePoint Site Target: {SHAREPOINT_SITE_NAME}/{SHAREPOINT_LIBRARY_NAME}" + (f"/{SHAREPOINT_FOLDER_PATH}" if SHAREPOINT_FOLDER_PATH else ""))
    print(f"FAISS Index: {FAISS_INDEX_PATH}, Metadata: {FAISS_METADATA_PATH}")
    print(f"To populate/update knowledgebase, POST to /update-knowledgebase")
    print(f"To chat, POST to /chat-with-sp-docs with {{'query': 'your question', 'history': 'chat history string'}}")
    print(f"To list indexed docs, GET /list-indexed-documents")
    print(f"Ensure MS Graph Token Cache '{MS_GRAPH_TOKEN_CACHE_FILE}' is valid (run generate_token_graph.py if needed).")
    print(f"Backend accessible at http://localhost:5001")

    app.run(host='0.0.0.0', port=5001, debug=True, use_reloader=False) # use_reloader=False often better with global state like models/indexes