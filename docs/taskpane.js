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
    return msalInstance;
  }
  
  try {
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
  if (!msalInstance) {
    await initializeMsal();
  }
  
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
    // Show downloading status
    statusDiv.className = "downloading";
    statusDiv.textContent = "🔐 Authenticating...";

    await initializeMsal();
    
    statusDiv.textContent = "📧 Fetching email content...";
    
    const accessToken = await getToken();

    const itemId = Office.context.mailbox.item.itemId;
    const graphItemId = encodeURIComponent(itemId);

    statusDiv.textContent = "⬇️ Downloading email...";

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

    const subject = Office.context.mailbox.item.subject || "email";
    const filename = subject.replace(/[/\\?%*:|"<>]/g, '-') + ".eml";

    const url = URL.createObjectURL(emlBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    statusDiv.className = "success";
    statusDiv.textContent = "✅ Email downloaded successfully!";
    
    // Auto-close the pane after 2 seconds
    setTimeout(() => {
      if (Office.context.ui) {
        Office.context.ui.closeContainer();
      }
    }, 2000);

  } catch (error) {
    statusDiv.className = "error";
    statusDiv.textContent = `❌ Error: ${error.message}`;
    console.error("Download error:", error);
  }
}

/* Initialize Office add-in and auto-start download */
Office.onReady((info) => {
  console.log("Office.onReady triggered", info);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("Host is Outlook, initializing auto-download...");
    
    // Hide sideload message, show app
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "block";

    // Auto-start download immediately
    downloadEmailAsEml();
  }
});