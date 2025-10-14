import * as msal from "@azure/msal-browser";

/* Azure AD MSAL config */
const msalConfig = {
  auth: {
    clientId: "10f65a22-c90e-44bc-9c3f-dbb90c8d6a92",
    redirectUri: window.location.origin
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
    // Create and initialize MSAL instance
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
  statusDiv.style.color = "blue";
  statusDiv.textContent = "Initializing authentication...";

  try {
    // Ensure MSAL is initialized before proceeding
    await initializeMsal();
    
    statusDiv.textContent = "Signing in and fetching email...";
    statusDiv.style.color = "green";
    
    const accessToken = await getToken();

    // Get the itemId and encode it for Graph API
    const itemId = Office.context.mailbox.item.itemId;
    const graphItemId = encodeURIComponent(itemId);

    statusDiv.textContent = "Downloading email content...";

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

    statusDiv.textContent = "Download started!";
  } catch (error) {
    statusDiv.style.color = "red";
    statusDiv.textContent = `Error: ${error.message}`;
    console.error("Download error:", error);
  }
}

/* Initialize Office add-in */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Hide the sideload message, show the app
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Attach click handler to the button
    document.getElementById("downloadBtn").onclick = downloadEmailAsEml;
    
    console.log("Office add-in ready");
  }
});