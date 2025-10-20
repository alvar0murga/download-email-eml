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

/* Start MSAL - Create and initialize instance */
async function initializeMsal() {
  if (msalInstance) {
    return msalInstance; // Already initialized
  }
  
  try {
    // Create and initialize MSAL instance using global msal object
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    console.log("MSAL Initialized successfully");
    return msalInstance;
  } catch (error) {
    console.error("Error initializing MSAL:", error);
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
    console.error("Login error:", error);
    throw error;
  }
}

/* Get access token silently or interactively */
async function getToken() {
  // Ensure MSAL is initialized first
  if (!msalInstance) {
    await initializeMsal();
  }
  
  // Now safely get accounts
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
    // Interaction required (consent or MFA)
    if (error instanceof msal.InteractionRequiredAuthError) {
      const response = await msalInstance.acquireTokenPopup(tokenRequest);
      return response.accessToken;
    } else {
      console.error("Token acquisition error:", error);
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
    console.log("Attempting direct MIME download...");
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 1)...";
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    if (response.ok) {
      const emlBlob = await response.blob();
      console.log("Direct MIME download successful");
      return emlBlob;
    } else {
      console.log("Direct MIME failed with status:", response.status);
    }
  } catch (error) {
    console.log("Direct MIME method failed:", error);
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 2: Try getting message details first, then MIME
  try {
    console.log("Attempting metadata + MIME download...");
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
      console.log("Got metadata for:", metadata.subject);
      
      // Now try MIME with the confirmed ID
      const mimeResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${metadata.id}/$value`, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "message/rfc822"
        }
      });

      if (mimeResponse.ok) {
        const emlBlob = await mimeResponse.blob();
        console.log("Metadata + MIME download successful");
        return emlBlob;
      }
    }
  } catch (error) {
    console.log("Metadata + MIME method failed:", error);
  }

  // Wait a bit before trying next method
  await delay(1000);

  // Method 3: Create EML from JSON data
  try {
    console.log("Attempting JSON to EML conversion...");
    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content (Method 3)...";
    
    const fullResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}?$select=subject,body,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime,internetMessageHeaders`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (fullResponse.ok) {
      const message = await fullResponse.json();
      console.log("Got full message data, creating EML...");
      
      const emlContent = createEmlFromJson(message);
      const emlBlob = new Blob([emlContent], { type: 'message/rfc822' });
      console.log("JSON to EML conversion successful");
      return emlBlob;
    }
  } catch (error) {
    console.log("JSON to EML method failed:", error);
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
  const statusDiv = document.getElementById("status");
  
  try {
    // Show downloading status with SED branding
    statusDiv.className = "downloading";
    statusDiv.textContent = "🔐 SED Email Downloader - Authenticating...";

    // Ensure MSAL is initialized before proceeding
    await initializeMsal();
    
    statusDiv.textContent = "📧 SED Email Downloader - Fetching email...";
    
    const accessToken = await getToken();

    // Get the itemId
    const itemId = Office.context.mailbox.item.itemId;
    console.log("Starting download for item ID:", itemId);

    // Try multiple download methods
    const emlBlob = await downloadEmailWithRetry(accessToken, itemId, statusDiv);

    // Clean subject for filename
    const subject = Office.context.mailbox.item.subject || "email";
    const filename = subject.replace(/[/\\?%*:|"<>]/g, '-') + ".eml";

    // Create a temporary download link and click it
    const url = URL.createObjectURL(emlBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    statusDiv.className = "success";
    statusDiv.textContent = "✅ SED Email Downloader - Download completed successfully!";
    
    // Auto-close the pane after 3 seconds
    setTimeout(() => {
      if (Office.context.ui) {
        Office.context.ui.closeContainer();
      }
    }, 3000);

  } catch (error) {
    statusDiv.className = "error";
    console.error("Download error:", error);
    
    if (error.message.includes("503")) {
      statusDiv.textContent = "❌ SED Email Downloader - Service temporarily unavailable. Please try again later.";
    } else if (error.message.includes("All download methods failed")) {
      statusDiv.textContent = "❌ SED Email Downloader - Unable to download email. Please check your connection and try again.";
    } else if (error.message.includes("404")) {
      statusDiv.textContent = "❌ SED Email Downloader - Email not found. Try refreshing Outlook.";
    } else if (error.message.includes("401") || error.message.includes("403")) {
      statusDiv.textContent = "❌ SED Email Downloader - Authentication error. Please try again.";
    } else {
      statusDiv.textContent = `❌ SED Email Downloader - Error: ${error.message}`;
    }
  }
}

/* Initialize Office add-in with SED branding */
console.log("SED Email Downloader - Script loaded, waiting for Office...");

// Add fallback initialization
document.addEventListener('DOMContentLoaded', function() {
  console.log("SED Email Downloader - DOM loaded");
  
  // If Office.onReady doesn't work, show the app anyway after a delay
  setTimeout(() => {
    if (document.getElementById("sideload-msg") && document.getElementById("sideload-msg").style.display !== "none") {
      console.log("SED Email Downloader - Office.onReady didn't trigger, showing app anyway");
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "block";
      
      // Try to attach click handler
      const btn = document.getElementById("downloadBtn");
      if (btn) {
        btn.onclick = downloadEmailAsEml;
      }
    }
  }, 3000);
});

Office.onReady((info) => {
  console.log("SED Email Downloader - Office.onReady triggered", info);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("SED Email Downloader - Host is Outlook, initializing...");
    
    // Hide the sideload message, show the app
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
      console.log("SED Email Downloader - Hid sideload message");
    }
    
    if (appBody) {
      appBody.style.display = "block";
      console.log("SED Email Downloader - Showed app body");
    }

    // Attach click handler to the button
    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) {
      downloadBtn.onclick = downloadEmailAsEml;
      console.log("SED Email Downloader - Attached click handler");
    } else {
      console.error("SED Email Downloader - Download button not found");
    }
    
    console.log("SED Email Downloader ready at:", window.location.href);
  } else {
    console.log("SED Email Downloader - Host is not Outlook:", info.host);
  }
});