// Remove the import line since we're loading MSAL via CDN

/* Azure AD MSAL config */
const msalConfig = {
  auth: {
    clientId: "10f65a22-c90e-44bc-9c3f-dbb90c8d6a92",
    redirectUri: "https://alvar0murga.github.io/download-email-eml/"
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
    scopes: ["Mail.Read"]
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
    scopes: ["Mail.Read"],
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

    // Get the itemId and encode it for Graph API
    const itemId = Office.context.mailbox.item.itemId;
    const graphItemId = encodeURIComponent(itemId);

    statusDiv.textContent = "⬇️ SED Email Downloader - Downloading content...";

    // Call Microsoft Graph API to get the MIME content of the email
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    if (!response.ok) {
      throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
    }

    const emlBlob = await response.blob();

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
    if (error.message.includes("503")) {
      statusDiv.textContent = "❌ SED Email Downloader - Service temporarily unavailable. Please try again.";
    } else if (error.message.includes("404")) {
      statusDiv.textContent = "❌ SED Email Downloader - Email not found. Try refreshing Outlook.";
    } else if (error.message.includes("401") || error.message.includes("403")) {
      statusDiv.textContent = "❌ SED Email Downloader - Authentication error. Please try again.";
    } else {
      statusDiv.textContent = `❌ SED Email Downloader - Error: ${error.message}`;
    }
    console.error("Download error:", error);
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