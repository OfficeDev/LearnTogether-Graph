//MSAL configuration
const msalConfig = {
  auth: {
    ...config,
    redirectUri: window.location.origin
  }
};
const msalRequest = { scopes: ['User.Read.All', 'Sites.Read.All', 'Calendars.Read', 'Mail.Read'] };

//Initialize MSAL client
const msalClient = new msal.PublicClientApplication(msalConfig);

// Log the user in
async function signIn() {
  const authResult = await msalClient.loginPopup(msalRequest);
  sessionStorage.setItem('msalAccount', authResult.account.username);
}
// try to sign user automatically in based on their previous session
async function silentSignIn() {
  const account = sessionStorage.getItem('msalAccount');
  if (!account) {
    return false;
  }

  const msalRequestSilent = {
    scopes: msalRequest.scopes,
    account: msalClient.getAccountByUsername(account)
  };
  try {
    const authResult = await msalClient.acquireTokenSilent(msalRequestSilent);
    sessionStorage.setItem('msalAccount', authResult.account.username);
    return true;
  }
  catch {
    return false;
  }
}
//Get token from Graph
async function getToken() {
  const account = sessionStorage.getItem('msalAccount');
  if (!account) {
    throw new Error(
      'User info cleared from session. Please sign out and sign in again.');
  }
  try {
    // First, attempt to get the token silently
    const silentRequest = {
      scopes: msalRequest.scopes,
      account: msalClient.getAccountByUsername(account)
    };

    const silentResult = await msalClient.acquireTokenSilent(silentRequest);
    return silentResult.accessToken;
  } catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
    if (silentError instanceof msal.InteractionRequiredAuthError) {
      const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
      return interactiveResult.accessToken;
    } else {
      throw silentError;
    }
  }
}

