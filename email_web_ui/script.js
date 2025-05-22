// script.js (from previous response - should still be mostly fine)
// Ensure this file is in your frontend directory (e.g., unified_email_ui/script.js)

document.addEventListener('DOMContentLoaded', () => {
    const emailListEl = document.getElementById('emailList');
    const refreshAllInboxBtn = document.getElementById('refreshAllInboxBtn');
    const statusMessagesGlobalDiv = document.getElementById('statusMessagesGlobal');
    
    const gmailAuthStatusEl = document.getElementById('gmailAuthStatus');
    const connectGmailBtn = document.getElementById('connectGmailBtn');
    const outlookAuthStatusEl = document.getElementById('outlookAuthStatus');

    // Modal elements
    const draftModal = $('#draftModal'); 
    const modalOriginalSender = document.getElementById('modalOriginalSender');
    const modalOriginalSubject = document.getElementById('modalOriginalSubject');
    const modalOriginalBody = document.getElementById('modalOriginalBody');
    const modalReplyTo = document.getElementById('modalReplyTo');
    const modalReplySubject = document.getElementById('modalReplySubject');
    const modalReplyBody = document.getElementById('modalReplyBody');
    const sendReplyBtnModal = document.getElementById('sendReplyBtnModal');
    const regenerateDraftBtnModal = document.getElementById('regenerateDraftBtnModal');
    const modalStatusMessages = document.getElementById('modalStatusMessages');

    const backendUrl = 'http://localhost:5000';
    const USER_NAME_FOR_PROMPT = "Khurram"; // Should match USER_NAME in app.py for consistency
    let currentEmailForReply = null; 

    function displayGlobalStatus(message, type = "info", persistent = false) {
        const alertId = `global-alert-${Date.now()}`;
        statusMessagesGlobalDiv.innerHTML = `<div id="${alertId}" class="alert alert-${type} alert-dismissible fade show" role="alert">
            ${escapeHtml(message)}
            <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        </div>`;
        if (!persistent && (type === "success" || type === "info")) {
            setTimeout(() => {
                const alertElement = document.getElementById(alertId);
                if (alertElement) $(alertElement).alert('close');
            }, 7000);
        }
    }
    
    function displayModalStatus(message, type = "info", persistent = false) {
        const alertId = `modal-alert-${Date.now()}`;
        modalStatusMessages.innerHTML = `<div id="${alertId}" class="alert alert-${type} alert-dismissible fade show" role="alert">
            ${escapeHtml(message)}
            <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        </div>`;
         if (!persistent && (type !== "danger" && type !== "warning")) {
            setTimeout(() => {
                const alertElement = document.getElementById(alertId);
                if (alertElement) $(alertElement).alert('close');
            }, 7000);
        }
    }

    async function checkAuthStatus() {
        try {
            const response = await fetch(`${backendUrl}/auth-status`);
            if (!response.ok) {
                const errText = await response.text();
                throw new Error(`Auth status check failed: ${response.status} - ${errText}`);
            }
            const status = await response.json();
            
            if (status.gmail) {
                gmailAuthStatusEl.textContent = "Gmail: Connected";
                gmailAuthStatusEl.className = "badge badge-success mr-2 p-2";
                connectGmailBtn.style.display = 'none';
            } else {
                gmailAuthStatusEl.textContent = "Gmail: Not Connected";
                gmailAuthStatusEl.className = "badge badge-danger mr-2 p-2";
                connectGmailBtn.style.display = 'inline-block';
            }

            if (status.outlook) {
                outlookAuthStatusEl.textContent = "Outlook: Connected";
                outlookAuthStatusEl.className = "badge badge-success mr-2 p-2";
            } else {
                outlookAuthStatusEl.textContent = "Outlook: Not Connected / Not on Windows";
                outlookAuthStatusEl.className = "badge badge-warning mr-2 p-2";
            }
        } catch (error) {
            console.error("Error checking auth status:", error);
            gmailAuthStatusEl.textContent = "Gmail: Error";
            outlookAuthStatusEl.textContent = "Outlook: Error";
            gmailAuthStatusEl.className = "badge badge-warning mr-2 p-2";
            outlookAuthStatusEl.className = "badge badge-warning mr-2 p-2";
            displayGlobalStatus(`Failed to check auth status: ${error.message}. Ensure backend is running.`, "danger", true);
        }
    }

    connectGmailBtn.addEventListener('click', async () => {
        displayGlobalStatus("Attempting to connect Gmail. A browser window might open on your server for authentication (if running Flask locally)...", "info", true);
        connectGmailBtn.disabled = true;
        try {
            const response = await fetch(`${backendUrl}/initiate-gmail-auth`);
            const data = await response.json(); 
            if (response.ok && data.status === 'success') {
                displayGlobalStatus(data.message + " Auth status will refresh shortly.", "success", true);
            } else {
                displayGlobalStatus(`Gmail connection failed: ${data.message || data.error || 'Unknown error from server'}`, "danger", true);
            }
        } catch (error) {
            displayGlobalStatus(`Error initiating Gmail connection: ${error.message}`, "danger", true);
        } finally {
            connectGmailBtn.disabled = false;
            setTimeout(checkAuthStatus, 3000); 
        }
    });

    async function fetchAndDisplayEmails() {
        displayGlobalStatus("Fetching all inbox emails...", "info");
        emailListEl.innerHTML = '<li class="list-group-item text-center">Loading... <div class="spinner-border spinner-border-sm"></div></li>';
        try {
            const response = await fetch(`${backendUrl}/emails?folder=inbox`);
            if (!response.ok) {
                const errData = await response.json().catch(() => ({error: "Server error fetching emails"}));
                throw new Error(`Failed to fetch emails: ${response.status} - ${errData.error}`);
            }
            const emails = await response.json();
            emailListEl.innerHTML = '';
            if (emails.length === 0) {
                emailListEl.innerHTML = '<li class="list-group-item">No emails in combined inbox.</li>';
                displayGlobalStatus("Inboxes are empty or no emails fetched.", "info");
                return;
            }
            emails.forEach(email => {
                const platformBadgeClass = email.platform === 'gmail' ? 'badge-danger' : 'badge-info';
                const platformName = email.platform === 'gmail' ? 'Gmail' : 'Outlook';
                const platformBadge = `<span class="badge ${platformBadgeClass} platform-badge">${platformName}</span>`;
                
                let dateStr = 'N/A';
                if (email.date) {
                    try {
                        // Attempt to parse, handling potential timezone offsets correctly if present
                        const dateObj = new Date(email.date);
                        dateStr = dateObj.toLocaleString();
                    } catch (e) {
                        dateStr = escapeHtml(email.date); // Show raw if parsing fails
                    }
                }

                const listItem = document.createElement('li');
                listItem.className = 'list-group-item email-item';
                listItem.innerHTML = `
                    <div class="d-flex w-100 justify-content-between">
                        <h5 class="mb-1">${platformBadge}${escapeHtml(email.subject || '(No Subject)')}</h5>
                        <small title="${escapeHtml(email.date)}">${dateStr}</small>
                    </div>
                    <p class="mb-1"><strong>From:</strong> ${escapeHtml(email.from || 'N/A')}</p>
                    <small class="email-snippet">${escapeHtml(email.snippet || '')}</small>
                    <button class="btn btn-sm btn-outline-primary float-right draft-reply-btn" 
                            data-platform="${email.platform}" 
                            data-id="${email.id}" 
                            data-thread-id="${email.threadId || ''}"
                            data-message-id-header="${escapeHtml(email.message_id_header || '')}"
                            data-references-header="${escapeHtml(email.references_header || '')}"
                            data-in-reply-to-header="${escapeHtml(email.in_reply_to_header || '')}"
                            >Draft Reply</button>
                `;
                emailListEl.appendChild(listItem);
            });
            displayGlobalStatus("Inboxes loaded successfully.", "success");
        } catch (error) {
            console.error('Error fetching emails:', error);
            emailListEl.innerHTML = `<li class="list-group-item list-group-item-danger">Error loading emails: ${error.message}</li>`;
            displayGlobalStatus(`Error fetching emails: ${error.message}`, "danger", true);
        }
    }

    emailListEl.addEventListener('click', async (event) => {
        const button = event.target.closest('.draft-reply-btn');
        if (button) {
            currentEmailForReply = { 
                platform: button.dataset.platform,
                id: button.dataset.id, 
                threadId: button.dataset.threadId,
                messageIdHeader: button.dataset.messageIdHeader,
                referencesHeader: button.dataset.referencesHeader,
                inReplyToHeader: button.dataset.inReplyToHeader,
                fullDetails: null 
            };

            modalStatusMessages.innerHTML = '';
            displayModalStatus("Loading email details and drafting reply...", "info");
            draftModal.modal('show');
            
            modalOriginalSender.textContent = 'Loading...';
            modalOriginalSubject.textContent = 'Loading...';
            modalOriginalBody.textContent = 'Loading...';
            modalReplyTo.value = '';
            modalReplySubject.value = '';
            modalReplyBody.value = 'AI is drafting... Please wait.';
            sendReplyBtnModal.disabled = true;
            regenerateDraftBtnModal.disabled = true;

            try {
                const detailsResponse = await fetch(`${backendUrl}/email-details?platform=${currentEmailForReply.platform}&id=${currentEmailForReply.id}`);
                if (!detailsResponse.ok) {
                    const errData = await detailsResponse.json().catch(() => ({error:"Error fetching email details from server."}));
                    // This is where the error from your screenshot likely originates if backend returns error JSON
                    throw new Error(`Could not fetch details: ${detailsResponse.status} - ${errData.error}`);
                }
                const emailDetails = await detailsResponse.json();
                
                if (!emailDetails || emailDetails.error) { 
                    throw new Error(emailDetails.error || "Email details response from backend was empty or contained an error.");
                }

                currentEmailForReply.fullDetails = emailDetails;

                modalOriginalSender.textContent = escapeHtml(emailDetails.from);
                modalOriginalSubject.textContent = escapeHtml(emailDetails.subject);
                modalOriginalBody.textContent = escapeHtml(emailDetails.body);

                modalReplyTo.value = extractEmailAddress(emailDetails.from);
                modalReplySubject.value = emailDetails.subject && emailDetails.subject.toLowerCase().startsWith("re:") ? emailDetails.subject : `Re: ${emailDetails.subject || ''}`;
                
                const draftResponse = await fetch(`${backendUrl}/draft-ai-reply`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        platform: currentEmailForReply.platform,
                        userName: USER_NAME_FOR_PROMPT, 
                        sender: emailDetails.from,
                        subject: emailDetails.subject,
                        body: emailDetails.body
                    })
                });
                const draftData = await draftResponse.json();
                if (draftResponse.ok && draftData.draft) {
                    modalReplyBody.value = draftData.draft;
                    displayModalStatus("AI draft generated!", "success");
                } else {
                    modalReplyBody.value = `Error generating draft: ${draftData.error || 'Unknown error'}`;
                    displayModalStatus(`Error from AI drafter: ${draftData.error || 'Unknown error'}`, "danger", true);
                }
            } catch (error) {
                console.error("Error in draft process:", error);
                modalReplyBody.value = `Error: ${error.message}`; // This will show the specific error from backend
                displayModalStatus(`Error during drafting process: ${error.message}`, "danger", true);
            } finally {
                sendReplyBtnModal.disabled = false;
                regenerateDraftBtnModal.disabled = false;
            }
        }
    });

    regenerateDraftBtnModal.addEventListener('click', async () => {
        if (currentEmailForReply && currentEmailForReply.fullDetails) {
            displayModalStatus("Regenerating AI draft...", "info");
            modalReplyBody.value = "AI is regenerating draft...";
            sendReplyBtnModal.disabled = true;
            regenerateDraftBtnModal.disabled = true;
            try {
                 const draftResponse = await fetch(`${backendUrl}/draft-ai-reply`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        platform: currentEmailForReply.platform,
                        userName: USER_NAME_FOR_PROMPT,
                        sender: currentEmailForReply.fullDetails.from,
                        subject: currentEmailForReply.fullDetails.subject,
                        body: currentEmailForReply.fullDetails.body
                    })
                });
                const draftData = await draftResponse.json();
                if (draftResponse.ok && draftData.draft) {
                    modalReplyBody.value = draftData.draft;
                    displayModalStatus("AI draft regenerated!", "success");
                } else {
                    modalReplyBody.value = `Error regenerating draft: ${draftData.error || 'Unknown error'}`;
                    displayModalStatus(`Error from AI drafter: ${draftData.error || 'Unknown error'}`, "danger", true);
                }
            } catch (error) {
                modalReplyBody.value = `Error regenerating draft: ${error.message}`;
                displayModalStatus(`Could not connect to AI drafter: ${error.message}`, "danger", true);
            } finally {
                 sendReplyBtnModal.disabled = false;
                 regenerateDraftBtnModal.disabled = false;
            }
        } else {
            displayModalStatus("Original email data not available to regenerate draft. Please close and retry.", "warning", true);
        }
    });

    sendReplyBtnModal.addEventListener('click', async () => {
        // ... (send reply logic from previous full script.js, unchanged for this fix) ...
        if (!currentEmailForReply || !currentEmailForReply.fullDetails) {
            displayModalStatus("Error: No email context for sending reply.", "danger", true);
            return;
        }
        const to = modalReplyTo.value.trim();
        const subject = modalReplySubject.value.trim();
        const body = modalReplyBody.value.trim();

        if (!to || !subject || !body) {
            displayModalStatus("To, Subject, and Body are required to send.", "warning", true);
            return;
        }
        displayModalStatus("Sending reply...", "info");
        sendReplyBtnModal.disabled = true;
        regenerateDraftBtnModal.disabled = true;
        try {
            const payload = {
                platform: currentEmailForReply.platform,
                originalMessageId: currentEmailForReply.id, 
                originalThreadId: currentEmailForReply.threadId, 
                to: to, subject: subject, body: body,
                inReplyToHeader: currentEmailForReply.fullDetails.in_reply_to_header,
                referencesHeader: currentEmailForReply.fullDetails.references_header
            };
            const response = await fetch(`${backendUrl}/send-platform-reply`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await response.json();
            if (response.ok && result.status === 'success') {
                displayModalStatus(`Reply sent successfully via ${currentEmailForReply.platform}! ${result.message || ''}`, "success");
                setTimeout(() => draftModal.modal('hide'), 3000);
                fetchAndDisplayEmails(); 
            } else {
                displayModalStatus(`Error sending reply: ${result.message || 'Unknown error from server'}`, "danger", true);
                sendReplyBtnModal.disabled = false;
                regenerateDraftBtnModal.disabled = false;
            }
        } catch (error) {
            console.error("Error sending reply:", error);
            displayModalStatus(`Failed to send reply: ${error.message}`, "danger", true);
            sendReplyBtnModal.disabled = false;
            regenerateDraftBtnModal.disabled = false;
        }
    });
    
    function extractEmailAddress(fromHeader) {
        if (!fromHeader) return "";
        const match = fromHeader.match(/([\w\.-]+@[\w\.-]+)/); 
        return match ? match[0] : fromHeader; 
    }

    function escapeHtml(unsafe) {
        if (typeof unsafe !== 'string') return '';
        return unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
    }
    
    // Initial actions
    checkAuthStatus(); 
    refreshAllInboxBtn.addEventListener('click', () => {
        fetchAndDisplayEmails(); // checkAuthStatus is not strictly needed before every refresh here
    });
    fetchAndDisplayEmails(); 
});