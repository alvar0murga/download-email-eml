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
    scopes: ["https://graph.microsoft.com/Mail.Read"]
  };

  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    console.log("Login successful:", loginResponse.account.username);
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
    scopes: ["https://graph.microsoft.com/Mail.Read"],
    account: account
  };

  try {
    const response = await msalInstance.acquireTokenSilent(tokenRequest);
    console.log("Token acquired successfully");
    return response.accessToken;
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      console.log("Interactive login required");
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
    statusDiv.className = "downloading";
    statusDiv.textContent = "🔐 Authenticating...";

    await initializeMsal();
    
    statusDiv.textContent = "📧 Fetching email content...";
    
    const accessToken = await getToken();
    console.log("Access token obtained, length:", accessToken.length);

    const itemId = Office.context.mailbox.item.itemId;
    console.log("Original Item ID:", itemId);
    
    const graphItemId = encodeURIComponent(itemId);
    console.log("Encoded Item ID:", graphItemId);

    statusDiv.textContent = "⬇️ Downloading email...";

    const graphUrl = `https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/$value`;
    console.log("Making request to:", graphUrl);

    const response = await fetch(graphUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "message/rfc822"
      }
    });

    console.log("Response status:", response.status);
    console.log("Response statusText:", response.statusText);
    
    // Log response headers
    for (let [key, value] of response.headers.entries()) {
      console.log(`Response header ${key}: ${value}`);
    }

    if (!response.ok) {
      let errorDetails = `${response.status} ${response.statusText}`;
      
      try {
        const errorBody = await response.text();
        console.log("Error response body:", errorBody);
        errorDetails += ` - ${errorBody}`;
      } catch (e) {
        console.log("Could not read error response body");
      }
      
      throw new Error(`Graph API error: ${errorDetails}`);
    }

    const emlBlob = await response.blob();
    console.log("Blob received, size:", emlBlob.size);

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
    
    setTimeout(() => {
      if (Office.context.ui) {
        Office.context.ui.closeContainer();
      }
    }, 2000);

  } catch (error) {
    statusDiv.className = "error";
    console.error("Full error object:", error);
    
    if (error.message.includes("503")) {
      statusDiv.textContent = "❌ Service temporarily unavailable. Please try again in a moment.";
    } else if (error.message.includes("404")) {
      statusDiv.textContent = "❌ Email not found. Try refreshing Outlook.";
    } else if (error.message.includes("401") || error.message.includes("403")) {
      statusDiv.textContent = "❌ Authentication error. Please try again.";
    } else {
      statusDiv.textContent = `❌ Error: ${error.message}`;
    }
  }
}

/* Initialize Office add-in and auto-start download */
Office.onReady((info) => {
  console.log("Office.onReady triggered", info);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("Host is Outlook, initializing auto-download...");
    console.log("Current location:", window.location.href);
    
    // Hide sideload message, show app
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "block";

    // Auto-start download immediately
    downloadEmailAsEml();
  }
});