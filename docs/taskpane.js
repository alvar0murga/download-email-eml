/* Azure AD MSAL config */
const msalConfig = {
  auth: {
    clientId: "10f65a22-c90e-44bc-9c3f-dbb90c8d6a92",
    redirectUri: "https://alvar0murga.github.io/download-email-eml/"
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  }
};

let msalInstance = null;
let isDownloading = false;

/* Start MSAL - Create and initialize instance */
async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }
  
  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    return msalInstance;
  } catch (error) {
    msalInstance = null;
    throw error;
  }
}

/* Sign in user with popup */
async function signIn() {
  if (!msalInstance) {
    throw new Error("MSAL not initialized");
  }
  
  const loginRequest = {
    scopes: ["https://graph.microsoft.com/Mail.Read"]
  };

  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    return loginResponse.account;
  } catch (error) {
    throw error;
  }
}

/* Get access token silently or interactively */
async function getToken() {
  if (!msalInstance) {
    await initializeMsal();
  }
  
  let account = msalInstance.getAllAccounts()[0];
  if (!account) {
    account = await signIn();
  }

  const tokenRequest = {
    scopes: ["https://graph.microsoft.com/Mail.Read"],
    account: account
  };

  try {
    const response = await msalInstance.acquireTokenSilent(tokenRequest);
    return response.accessToken;
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      const response = await msalInstance.acquireTokenPopup(tokenRequest);
      return response.accessToken;
    } else {
      throw error;
    }
  }
}

/* Helper function to wait/delay */
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/* Try multiple approaches to download email with detailed error reporting */
async function downloadEmailWithRetry(accessToken, itemId, statusDiv) {
  let errorDetails = [];
  
  // Simple validation
  if (!itemId) {
    throw new Error("Item ID is null or undefined");
  }
  
  statusDiv.textContent = `⬇️ Processing item ID: ${itemId.substring(0, 30)}...`;
  
  // Try different encoding approaches for the item ID
  const encodingMethods = [
    { name: "Direct (no encoding)", value: itemId },
    { name: "URI Component", value: encodeURIComponent(itemId) },
    { name: "URI", value: encodeURI(itemId) },
    { name: "Base64 safe", value: btoa(itemId).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '') }
  ];
  
  for (let i = 0; i < encodingMethods.length; i++) {
    const method = encodingMethods[i];
    const graphItemId = method.value;
    
    try {
      statusDiv.textContent = `⬇️ Trying ${method.name} encoding...`;
      
      // Try Method 3 first (JSON to EML) as it's most likely to work
      const fullResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}?$select=subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "application/json"
        }
      });

      if (fullResponse.ok) {
        const message = await fullResponse.json();
        statusDiv.textContent = `✅ Success with ${method.name} encoding!`;
        
        const emlContent = createEmlFromJson(message);
        const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
        return emlBlob;
      } else {
        const errorText = await fullResponse.text();
        errorDetails.push(`${method.name} failed: ${fullResponse.status} ${fullResponse.statusText} - ${errorText}`);
      }
    } catch (error) {
      errorDetails.push(`${method.name} error: ${error.message}`);
    }
    
    await delay(500);
  }
  
  // Special handling for very recent emails - try to find by subject
  try {
    statusDiv.textContent = "⬇️ Trying recent email fallback method...";
    
    // Get the subject from the Outlook context
    const currentSubject = Office.context.mailbox.item.subject;
    const currentFrom = Office.context.mailbox.item.from?.emailAddress?.address;
    
    if (currentSubject && currentFrom) {
      // Search for the email by subject and sender in recent messages
      const searchQuery = `subject:"${currentSubject}" AND from:${currentFrom}`;
      const encodedQuery = encodeURIComponent(searchQuery);
      
      const searchResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$search="${encodedQuery}"&$top=5&$select=id,subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "application/json"
        }
      });
      
      if (searchResponse.ok) {
        const searchResults = await searchResponse.json();
        
        if (searchResults.value && searchResults.value.length > 0) {
          // Find the most recent match
          const matchedEmail = searchResults.value[0];
          statusDiv.textContent = `✅ Found recent email via search!`;
          
          const emlContent = createEmlFromJson(matchedEmail);
          const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
          return emlBlob;
        } else {
          errorDetails.push("Recent email fallback: No matching emails found in search");
        }
      } else {
        const errorText = await searchResponse.text();
        errorDetails.push(`Recent email fallback failed: ${searchResponse.status} - ${errorText}`);
      }
    } else {
      errorDetails.push("Recent email fallback: Missing subject or sender information");
    }
  } catch (error) {
    errorDetails.push(`Recent email fallback error: ${error.message}`);
  }
  
  // If all methods failed, show detailed error with helpful message
  const detailedError = `All methods failed for item ID: ${itemId.substring(0, 50)}...\n${errorDetails.join('\n')}\n\n🕐 This appears to be a very recent email (received within the last few hours).\n📧 Recent emails sometimes take time to fully sync with Microsoft Graph API.\n\n💡 Solutions:\n- Wait 30-60 minutes and try again\n- Try refreshing the email in Outlook\n- The email should work normally once it's fully synced`;
  throw new Error(detailedError);
}

/* Create EML format from JSON message data */
function createEmlFromJson(message) {
  const date = new Date(message.receivedDateTime).toUTCString();
  const from = message.sender?.emailAddress?.address || "unknown@unknown.com";
  const fromName = message.sender?.emailAddress?.name || "";
  const to = message.toRecipients?.map(r => `${r.emailAddress.name || ""} <${r.emailAddress.address}>`).join(', ') || "";
  const cc = message.ccRecipients?.map(r => `${r.emailAddress.name || ""} <${r.emailAddress.address}>`).join(', ') || "";
  const bcc = message.bccRecipients?.map(r => `${r.emailAddress.name || ""} <${r.emailAddress.address}>`).join(', ') || "";
  const subject = message.subject || "(No Subject)";
  const body = message.body?.content || "";
  const contentType = message.body?.contentType === "html" ? "text/html" : "text/plain";

  let eml = `Date: ${date}\r\n`;
  eml += `From: ${fromName ? `${fromName} <${from}>` : from}\r\n`;
  if (to) eml += `To: ${to}\r\n`;
  if (cc) eml += `Cc: ${cc}\r\n`;
  if (bcc) eml += `Bcc: ${bcc}\r\n`;
  eml += `Subject: ${subject}\r\n`;
  eml += `MIME-Version: 1.0\r\n`;
  eml += `Content-Type: ${contentType}; charset=utf-8\r\n`;
  eml += `Content-Transfer-Encoding: 8bit\r\n`;
  eml += `\r\n`;
  eml += body;

  return eml;
}

/* Download using a different method to avoid search bar issue */
function triggerDownload(blob, filename) {
  // Create download link with proper MIME type
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  
  // Force the download by making it invisible and clicking it
  link.style.display = 'none';
  document.body.appendChild(link);
  
  // Try to trigger download
  try {
    link.click();
  } catch (error) {
    // Fallback: show a visible download link
    link.style.display = 'block';
    link.style.color = '#0078d4';
    link.style.textDecoration = 'underline';
    link.style.marginTop = '10px';
    link.textContent = 'Click here to download your email file';
    
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
      statusDiv.appendChild(document.createElement('br'));
      statusDiv.appendChild(link);
    }
  }
  
  // Clean up after download
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 5000);
}

/* Download the currently selected email as .eml */
async function downloadEmailAsEml() {
  if (isDownloading) {
    return;
  }
  
  isDownloading = true;
  const statusDiv = document.getElementById("status");
  const downloadBtn = document.getElementById("downloadBtn");
  
  try {
    // Disable button and show progress
    if (downloadBtn) {
      downloadBtn.disabled = true;
      downloadBtn.textContent = "⏳ Downloading...";
    }
    
    // Show downloading status
    if (statusDiv) {
      statusDiv.className = "downloading";
      statusDiv.textContent = "🔐 SED Email Downloader - Authenticating...";
    }

    await initializeMsal();
    
    if (statusDiv) {
      statusDiv.textContent = "📧 SED Email Downloader - Fetching email...";
    }
    
    const accessToken = await getToken();
    
    const itemId = Office.context.mailbox.item.itemId;

    if (!itemId) {
      throw new Error("No item ID found - make sure you're viewing an email");
    }

    const emlBlob = await downloadEmailWithRetry(accessToken, itemId, statusDiv);

    // Clean subject for filename and ensure .eml extension
    const subject = Office.context.mailbox.item.subject || "email";
    let filename = subject.replace(/[/\\?%*:|"<>]/g, '-');
    
    // Ensure the filename ends with .eml
    if (!filename.toLowerCase().endsWith('.eml')) {
      filename += '.eml';
    }

    if (statusDiv) {
      statusDiv.textContent = "💾 SED Email Downloader - Starting download...";
    }

    // Use the improved download method
    triggerDownload(emlBlob, filename);

    if (statusDiv) {
      statusDiv.className = "success";
      statusDiv.textContent = `✅ SED Email Downloader - Download completed! File: ${filename}`;
    }
    
    // Reset button
    if (downloadBtn) {
      downloadBtn.disabled = false;
      downloadBtn.textContent = "📧 Download Another Email";
    }

  } catch (error) {
    if (statusDiv) {
      statusDiv.className = "error";
      statusDiv.style.whiteSpace = "pre-wrap";
      statusDiv.style.fontSize = "12px";
      statusDiv.style.textAlign = "left";
      statusDiv.textContent = `❌ SED Email Downloader - Error Details:\n${error.message}`;
    }
    
    // Re-enable button for retry
    if (downloadBtn) {
      downloadBtn.disabled = false;
      downloadBtn.textContent = "📧 Try Download Again";
    }
  }
  
  isDownloading = false;
}

/* Initialize Office add-in */
document.addEventListener('DOMContentLoaded', function() {
  const downloadBtn = document.getElementById("downloadBtn");
  if (downloadBtn) {
    downloadBtn.onclick = downloadEmailAsEml;
  }
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Hide the sideload message, show the app
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    
    if (appBody) {
      appBody.style.display = "block";
    }

    // Set up button handler
    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) {
      downloadBtn.onclick = downloadEmailAsEml;
    }
    
    // AUTO-START THE DOWNLOAD after a short delay
    setTimeout(() => {
      downloadEmailAsEml();
    }, 2000);
  }
});