// Remove the import line since we're loading MSAL via CDN

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

// Don't create instance immediately - wait for initialization
let msalInstance = null;
let isDownloading = false;

function updateDebugInfo(key, value) {
  const element = document.getElementById(key);
  if (element) {
    element.textContent = value;
  }
}

/* Start MSAL - Create and initialize instance */
async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }
  
  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    updateDebugInfo('msal-ready', 'Yes');
    return msalInstance;
  } catch (error) {
    updateDebugInfo('msal-ready', 'Error: ' + error.message);
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

/* Try multiple approaches to download email */
async function downloadEmailWithRetry(accessToken, itemId, statusDiv) {
  const graphItemId = encodeURIComponent(itemId);
  
  // Method 1: Try direct MIME download
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 1)...";
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    if (response.ok) {
      const emlBlob = await response.blob();
      return emlBlob;
    }
  } catch (error) {
    // Continue to next method
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 2: Try getting message details first, then MIME
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 2)...";
    
    // First get the message metadata
    const metadataResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (metadataResponse.ok) {
      const metadata = await metadataResponse.json();
      
      // Now try MIME with the confirmed ID
      const mimeResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${metadata.id}/$value`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "message/rfc822"
        }
      });

      if (mimeResponse.ok) {
        const emlBlob = await mimeResponse.blob();
        return emlBlob;
      }
    }
  } catch (error) {
    // Continue to next method
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 3: Create EML from JSON data
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 3)...";
    
    const fullResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}?$select=subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (fullResponse.ok) {
      const message = await fullResponse.json();
      
      const emlContent = createEmlFromJson(message);
      const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
      return emlBlob;
    }
  } catch (error) {
    // All methods failed
  }

  throw new Error("All download methods failed");
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
    
    // Show downloading status with SED branding
    if (statusDiv) {
      statusDiv.className = "downloading";
      statusDiv.textContent = "🔐 SED Email Downloader - Authenticating...";
    }

    // Ensure MSAL is initialized before proceeding
    await initializeMsal();
    
    if (statusDiv) {
      statusDiv.textContent = "📧 SED Email Downloader - Fetching email...";
    }
    
    const accessToken = await getToken();

    // Get the itemId
    const itemId = Office.context.mailbox.item.itemId;

    if (!itemId) {
      throw new Error("No item ID found - make sure you're viewing an email");
    }

    // Try multiple download methods
    const emlBlob = await downloadEmailWithRetry(accessToken, itemId, statusDiv);

    // Clean subject for filename
    const subject = Office.context.mailbox.item.subject || "email";
    const filename = subject.replace(/[/\\?%*:|"<>]/g, '-') + ".eml";

    // Use a different download approach to avoid triggering search
    if (statusDiv) {
      statusDiv.textContent = "💾 SED Email Downloader - Preparing download...";
    }

    // Method 1: Try using URL.createObjectURL with a different approach
    try {
      const url = URL.createObjectURL(emlBlob);
      
      // Create download link but don't append to body immediately
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      
      // Use mousedown event instead of click to avoid Outlook interference
      const mouseEvent = new MouseEvent('mousedown', {
        view: window,
        bubbles: true,
        cancelable: true
      });
      
      // Temporarily add to a hidden container
      const hiddenDiv = document.createElement('div');
      hiddenDiv.style.position = 'absolute';
      hiddenDiv.style.left = '-9999px';
      hiddenDiv.style.top = '-9999px';
      hiddenDiv.appendChild(a);
      document.body.appendChild(hiddenDiv);
      
      // Trigger download
      a.dispatchEvent(mouseEvent);
      a.click();
      
      // Clean up immediately
      setTimeout(() => {
        document.body.removeChild(hiddenDiv);
        URL.revokeObjectURL(url);
      }, 100);
      
    } catch (downloadError) {
      // Fallback: Show the blob URL for manual download
      const url = URL.createObjectURL(emlBlob);
      
      if (statusDiv) {
        statusDiv.className = "success";
        statusDiv.innerHTML = `✅ Email ready! <a href="${url}" download="${filename}" style="color: #0078d4; text-decoration: underline;">Click here to download</a>`;
      }
      
      // Clean up URL after 5 minutes
      setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 300000);
      
      isDownloading = false;
      if (downloadBtn) {
        downloadBtn.disabled = false;
        downloadBtn.textContent = "📧 Download Another Email";
        downloadBtn.onclick = downloadEmailAsEml;
      }
      return;
    }

    if (statusDiv) {
      statusDiv.className = "success";
      statusDiv.textContent = "✅ SED Email Downloader - Download completed! You can close this panel.";
    }
    
    // Change button to close button
    if (downloadBtn) {
      downloadBtn.disabled = false;
      downloadBtn.textContent = "✖️ Close Panel";
      downloadBtn.onclick = function() {
        // Just hide the panel content
        const appBody = document.getElementById("app-body");
        if (appBody) {
          appBody.style.display = "none";
        }
        document.body.innerHTML = '<div style="padding: 20px; text-align: center; font-family: Segoe UI; color: #107c10;">✅ Download completed! You can manually close this panel.</div>';
      };
    }

  } catch (error) {
    if (statusDiv) {
      statusDiv.className = "error";
      statusDiv.textContent = `❌ SED Email Downloader - Error: ${error.message}`;
    }
    
    // Re-enable button for retry
    if (downloadBtn) {
      downloadBtn.disabled = false;
      downloadBtn.textContent = "📧 Try Download Again";
      downloadBtn.onclick = downloadEmailAsEml;
    }
  }
  
  isDownloading = false;
}

// Keep all the other functions unchanged (initializeMsal, signIn, getToken, etc.)
/* Start MSAL - Create and initialize instance */
async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }
  
  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    updateDebugInfo('msal-ready', 'Yes');
    return msalInstance;
  } catch (error) {
    updateDebugInfo('msal-ready', 'Error: ' + error.message);
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

/* Try multiple approaches to download email */
async function downloadEmailWithRetry(accessToken, itemId, statusDiv) {
  const graphItemId = encodeURIComponent(itemId);
  
  // Method 1: Try direct MIME download
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 1)...";
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    if (response.ok) {
      const emlBlob = await response.blob();
      return emlBlob;
    }
  } catch (error) {
    // Continue to next method
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 2: Try getting message details first, then MIME
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 2)...";
    
    // First get the message metadata
    const metadataResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (metadataResponse.ok) {
      const metadata = await metadataResponse.json();
      
      // Now try MIME with the confirmed ID
      const mimeResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${metadata.id}/$value`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "message/rfc822"
        }
      });

      if (mimeResponse.ok) {
        const emlBlob = await mimeResponse.blob();
        return emlBlob;
      }
    }
  } catch (error) {
    // Continue to next method
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 3: Create EML from JSON data
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 3)...";
    
    const fullResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}?$select=subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (fullResponse.ok) {
      const message = await fullResponse.json();
      
      const emlContent = createEmlFromJson(message);
      const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
      return emlBlob;
    }
  } catch (error) {
    // All methods failed
  }

  throw new Error("All download methods failed");
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

/* Initialize Office add-in with SED branding */

// Add fallback initialization
document.addEventListener('DOMContentLoaded', function() {
  // Set up manual button click handler
  const downloadBtn = document.getElementById("downloadBtn");
  if (downloadBtn) {
    downloadBtn.onclick = downloadEmailAsEml;
  }
});

Office.onReady((info) => {
  updateDebugInfo('office-ready', 'Yes');
  
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

    // Set up manual button click handler
    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) {
      downloadBtn.onclick = downloadEmailAsEml;
    }
    
    // AUTO-START THE DOWNLOAD after a short delay
    setTimeout(() => {
      updateDebugInfo('auto-download', 'Starting...');
      downloadEmailAsEml();
    }, 2000);
  }
});