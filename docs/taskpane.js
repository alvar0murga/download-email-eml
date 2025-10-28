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

/* Try multiple approaches to download email */
async function downloadEmailWithRetry(accessToken, itemId, statusDiv) {
  const graphItemId = encodeURIComponent(itemId);
  
  // Method 1: Try direct MIME download
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content...";
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    if (response.ok) {
      const emlBlob = await response.blob();
      return emlBlob;
    } else {
      console.log("Method 1 failed with status:", response.status);
      const errorText = await response.text();
      console.log("Error response:", errorText);
    }
  } catch (error) {
    console.log("Method 1 error:", error);
  }

  await delay(1000);

  // Method 2: Try getting message details first, then MIME
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Fetching email data...";
    
    const metadataResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (metadataResponse.ok) {
      const metadata = await metadataResponse.json();
      console.log("Got metadata for:", metadata.subject);
      
      const mimeResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${metadata.id}/$value`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "message/rfc822"
        }
      });

      if (mimeResponse.ok) {
        const emlBlob = await mimeResponse.blob();
        return emlBlob;
      } else {
        console.log("Method 2 MIME failed with status:", mimeResponse.status);
      }
    } else {
      console.log("Method 2 metadata failed with status:", metadataResponse.status);
    }
  } catch (error) {
    console.log("Method 2 error:", error);
  }

  await delay(1000);

  // Method 3: Create EML from JSON data
  try {
    statusDiv.textContent = "⬇️ SED Email Downloader - Creating email file...";
    
    const fullResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}?$select=subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (fullResponse.ok) {
      const message = await fullResponse.json();
      console.log("Got full message data for:", message.subject);
      
      const emlContent = createEmlFromJson(message);
      const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
      return emlBlob;
    } else {
      console.log("Method 3 failed with status:", fullResponse.status);
      const errorText = await fullResponse.text();
      console.log("Method 3 error response:", errorText);
    }
  } catch (error) {
    console.log("Method 3 error:", error);
  }

  throw new Error("All download methods failed - check console for details");
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
  const url = URL.createObjectURL(blob);
  
  // Instead of using click(), use window.open which is less likely to trigger search
  const newWindow = window.open(url, '_blank');
  if (newWindow) {
    setTimeout(() => {
      newWindow.close();
      URL.revokeObjectURL(url);
    }, 1000);
  } else {
    // Fallback: create a link that user can click manually
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.textContent = 'Click here to download your email';
    link.style.color = '#0078d4';
    link.style.textDecoration = 'underline';
    link.style.display = 'block';
    link.style.marginTop = '10px';
    
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
      statusDiv.appendChild(document.createElement('br'));
      statusDiv.appendChild(link);
    }
    
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 300000); // Clean up after 5 minutes
  }
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

    console.log("Starting download process...");
    await initializeMsal();
    console.log("MSAL initialized");
    
    if (statusDiv) {
      statusDiv.textContent = "📧 SED Email Downloader - Fetching email...";
    }
    
    const accessToken = await getToken();
    console.log("Got access token, length:", accessToken ? accessToken.length : 0);
    
    const itemId = Office.context.mailbox.item.itemId;
    console.log("Item ID:", itemId);

    if (!itemId) {
      throw new Error("No item ID found - make sure you're viewing an email");
    }

    const emlBlob = await downloadEmailWithRetry(accessToken, itemId, statusDiv);
    console.log("Download successful, blob size:", emlBlob.size);

    // Clean subject for filename
    const subject = Office.context.mailbox.item.subject || "email";
    const filename = subject.replace(/[/\\?%*:|"<>]/g, '-') + ".eml";

    if (statusDiv) {
      statusDiv.textContent = "💾 SED Email Downloader - Starting download...";
    }

    // Use the new download method
    triggerDownload(emlBlob, filename);

    if (statusDiv) {
      statusDiv.className = "success";
      statusDiv.textContent = "✅ SED Email Downloader - Download completed! You can close this panel.";
    }
    
    // Reset button
    if (downloadBtn) {
      downloadBtn.disabled = false;
      downloadBtn.textContent = "📧 Download Another Email";
      downloadBtn.onclick = downloadEmailAsEml;
    }

  } catch (error) {
    console.error("Download error:", error);
    
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