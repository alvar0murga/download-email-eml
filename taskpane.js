import * as msal from "@azure/msal-browser";

/* Azure AD MSAL config */
const msalConfig = {
  auth: {
    clientId: "10f65a22-c90e-44bc-9c3f-dbb90c8d6a92", // Tu client ID
    redirectUri: "https://localhost" // Tu URI de redirección registrada en Azure
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

/* Sign in user with popup */
async function signIn() {
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
  statusDiv.style.color = "green";
  statusDiv.textContent = "Signing in and fetching email...";

  try {
    const accessToken = await getToken();

    // Get the itemId and encode it for Graph API
    const itemId = Office.context.mailbox.item.itemId;
    const graphItemId = encodeURIComponent(itemId);

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
    console.error(error);
  }
}

async function initializeMsal() {
  // Inicialización de MSAL
  await msalInstance.initialize();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Hide the sideload message, show the app
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Attach click handler to the button
    document.getElementById("downloadBtn").onclick = downloadEmailAsEml;
  }
});

