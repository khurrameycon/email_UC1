document.addEventListener('DOMContentLoaded', () => {
    const userInput = document.getElementById('userInput');
    const sendBtn = document.getElementById('sendBtn');
    const chatMessagesDiv = document.getElementById('chatMessages');
    const updateKbBtn = document.getElementById('updateKbBtn');
    const connectMsGraphBtn = document.getElementById('connectMsGraphBtn'); // For manual auth trigger
    const kbStatusDiv = document.getElementById('kbStatus');
    const indexedDocsListEl = document.getElementById('indexedDocsList');

    const backendChatUrl = 'http://localhost:5001'; // Ensure this matches your app_chat.py port
    let chatHistory = []; // Array of {role: 'user'/'assistant', content: '...'}

    function displayKbStatus(message, type = "info") {
        kbStatusDiv.textContent = message;
        kbStatusDiv.className = `mb-2 status-${type}`; // Simple status styling
    }

    function addMessageToChat(sender, message, sources = []) {
        const messageDiv = document.createElement('div');
        messageDiv.classList.add('message', sender === 'user' ? 'user-message' : 'ai-message');
        
        let messageHtml = `<div class="message-content">${escapeHtml(message).replace(/\n/g, '<br>')}</div>`;
        
        if (sources && sources.length > 0) {
            let sourcesHtml = '<div class="sources"><strong>Sources:</strong><ul>';
            sources.forEach(source => {
                sourcesHtml += `<li>${escapeHtml(source.name)} ${source.webUrl ? `(<a href="${source.webUrl}" target="_blank">link</a>)` : ''}</li>`;
            });
            sourcesHtml += '</ul></div>';
            messageHtml += sourcesHtml;
        }
        
        messageDiv.innerHTML = messageHtml;
        chatMessagesDiv.appendChild(messageDiv);
        chatMessagesDiv.scrollTop = chatMessagesDiv.scrollHeight; // Scroll to bottom
    }

    async function fetchAndDisplayIndexedDocs() {
        try {
            const response = await fetch(`${backendChatUrl}/list-indexed-documents`);
            if (!response.ok) {
                const errData = await response.json().catch(() => ({ error: "Could not fetch indexed docs list" }));
                displayKbStatus(`Error listing docs: ${errData.error}`, "error");
                return;
            }
            const data = await response.json();
            indexedDocsListEl.innerHTML = ''; // Clear
            if (data.documents && data.documents.length > 0) {
                data.documents.forEach(doc => {
                    const li = document.createElement('li');
                    li.className = 'list-group-item';
                    li.textContent = doc.name;
                    if (doc.webUrl) {
                        const link = document.createElement('a');
                        link.href = doc.webUrl;
                        link.target = '_blank';
                        link.textContent = ' (link)';
                        link.style.fontSize = '0.8em';
                        li.appendChild(link);
                    }
                    indexedDocsListEl.appendChild(li);
                });
                displayKbStatus(`Knowledge base loaded with ${data.documents.length} unique source documents.`, "ok");
            } else if(data.error) {
                displayKbStatus(data.error, "error");
                indexedDocsListEl.innerHTML = '<li class="list-group-item">Knowledge base not loaded.</li>';
            } else {
                indexedDocsListEl.innerHTML = '<li class="list-group-item">No documents currently indexed.</li>';
                displayKbStatus("Knowledge base is empty. Please update.", "info");
            }

        } catch (error) {
            console.error("Error fetching indexed documents:", error);
            displayKbStatus(`Error fetching indexed docs list: ${error.message}`, "error");
        }
    }

    updateKbBtn.addEventListener('click', async () => {
        displayKbStatus("Updating knowledge base... This may take several minutes depending on the number of documents.", "processing");
        updateKbBtn.disabled = true;
        connectMsGraphBtn.disabled = true;
        try {
            const response = await fetch(`${backendChatUrl}/update-knowledgebase`, { method: 'POST' });
            const result = await response.json();
            if (response.ok) {
                displayKbStatus(result.message || "Knowledge base update completed.", "ok");
                fetchAndDisplayIndexedDocs(); // Refresh the list
            } else {
                displayKbStatus(`Error updating knowledge base: ${result.error || 'Unknown error'}`, "error");
            }
        } catch (error) {
            console.error('Error updating knowledge base:', error);
            displayKbStatus(`Failed to update knowledge base: ${error.message}`, "error");
        } finally {
            updateKbBtn.disabled = false;
            connectMsGraphBtn.disabled = false;
        }
    });

    connectMsGraphBtn.addEventListener('click', async () => {
        // This button is for users to manually trigger the MS Graph auth device flow
        // if the backend indicates it's needed (e.g., token expired and silent failed)
        // The backend /get_msgraph_token_for_chat tries silent first.
        // For a better UX, the backend could have dedicated /login and /callback routes for MSAL web flow.
        // This button simulates initiating that if needed locally for a single user.
        displayKbStatus("Attempting to initiate Microsoft Graph connection... Follow backend console instructions if device flow starts.", "info");
        try {
            // A bit of a hack: The backend's get_msgraph_token_for_chat will log if interactive auth is needed.
            // This button doesn't directly *call* an auth URL, it's more of a reminder that the backend
            // might need it. A better way is for backend to return a 401 with auth_url if needed.
            // For now, this button is more of a placeholder to remind about the auth process.
            // Let's call an endpoint that might trigger device flow if token is bad.
            const response = await fetch(`${backendChatUrl}/update-knowledgebase`, { method: 'POST' }); // This endpoint requires token
             if (response.status === 401) { // Unauthorized
                displayKbStatus("Microsoft authentication required. Check backend console for device flow instructions if prompted, or ensure token cache is valid.", "warning");
            } else if (response.ok) {
                const result = await response.json();
                 displayKbStatus(result.message || "Knowledge base update process completed/started.", "ok");
                 fetchAndDisplayIndexedDocs();
            } else {
                const result = await response.json().catch(()=>({error: "Auth check failed"}));
                displayKbStatus(`Microsoft auth check: ${result.error || response.statusText }`, "warning");
            }

        } catch (error) {
            displayKbStatus(`Error with Microsoft connection: ${error.message}`, "error");
        }
    });


    async function handleUserMessage() {
        const query = userInput.value.trim();
        if (!query) return;

        addMessageToChat('user', query);
        userInput.value = ""; // Clear input
        sendBtn.disabled = true;

        // Construct history string for backend (simple concatenation)
        let historyString = "";
        chatHistory.forEach(msg => {
            historyString += `${msg.role === 'user' ? 'User' : 'AI'}: ${msg.content}\n`;
        });

        try {
            const response = await fetch(`${backendChatUrl}/chat-with-sp-docs`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query: query, history: historyString })
            });
            const data = await response.json();
            if (response.ok && data.response) {
                addMessageToChat('ai', data.response, data.sources);
                chatHistory.push({ role: 'user', content: query });
                chatHistory.push({ role: 'ai', content: data.response });
                // Keep history to a reasonable length (e.g., last 10 turns = 20 messages)
                if(chatHistory.length > 20) chatHistory.splice(0, chatHistory.length - 20);

            } else {
                addMessageToChat('ai', `Error: ${data.error || 'Could not get a response.'}`);
            }
        } catch (error) {
            console.error('Error during chat:', error);
            addMessageToChat('ai', `Error: ${error.message}`);
        } finally {
            sendBtn.disabled = false;
            userInput.focus();
        }
    }

    sendBtn.addEventListener('click', handleUserMessage);
    userInput.addEventListener('keypress', (event) => {
        if (event.key === 'Enter') {
            handleUserMessage();
        }
    });

    function escapeHtml(unsafe) {
        if (typeof unsafe !== 'string') return '';
        return unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
    }

    // Initial load
    fetchAndDisplayIndexedDocs(); // Load list of currently indexed documents on page start
});