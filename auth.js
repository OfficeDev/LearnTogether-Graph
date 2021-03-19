//MSAL configuration
const msalConfig = {
  auth: {
    ...config,
    redirectUri: 'http://localhost:8080'
  }
};
const msalRequest = { scopes: [] };
function ensureScope(scopes) {
  const scopesArray = Array.isArray(scopes) ? scopes : [scopes];
  scopesArray.forEach(scope => {
    if (!msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())) {
      msalRequest.scopes.push(scope);
    }
  })
}
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
    return;
  }

  const msalRequest = {
    scopes: [],
    account: msalClient.getAccountByUsername(account)
  };
  try {
    const authResult = await msalClient.acquireTokenSilent(msalRequest);
    sessionStorage.setItem('msalAccount', authResult.account.username);
  }
  catch { }
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

