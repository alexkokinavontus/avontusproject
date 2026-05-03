// Microsoft Entra ID (MSAL) Authentication
const MSAL_CONFIG = {
  auth: {
    clientId: "3977e66a-cdf1-419d-9d0d-70e8cf3a76ed",
    authority: "https://login.microsoftonline.com/bd98204b-b981-4d03-8796-356d537927eb",
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

const LOGIN_SCOPES = { scopes: ["User.Read", "openid", "profile", "email"] };
let msalInstance = null;

async function getMsal() {
  if (msalInstance) return msalInstance;
  const { PublicClientApplication } = await import("@azure/msal-browser");
  msalInstance = new PublicClientApplication(MSAL_CONFIG);
  await msalInstance.initialize();
  return msalInstance;
}

function buildUser(acct) {
  return {
    name: acct.name || acct.username,
    email: acct.username,
    tenantId: acct.tenantId,
    initials: (acct.name || acct.username || "?").split(" ").map(w => w[0]).join("").slice(0, 2).toUpperCase(),
  };
}

export async function getStoredUser() {
  try {
    const msal = await getMsal();
    await msal.handleRedirectPromise();
    const accounts = msal.getAllAccounts();
    if (!accounts.length) return null;
    msal.setActiveAccount(accounts[0]);
    return buildUser(accounts[0]);
  } catch { return null; }
}

export async function signIn() {
  const msal = await getMsal();
  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    msal.setActiveAccount(accounts[0]);
    return buildUser(accounts[0]);
  }
  try {
    const result = await msal.loginPopup(LOGIN_SCOPES);
    msal.setActiveAccount(result.account);
    return buildUser(result.account);
  } catch (e) {
    if (e.errorCode === "popup_window_error" || e.errorCode === "empty_window_error") {
      await msal.loginRedirect(LOGIN_SCOPES);
      return null;
    }
    throw e;
  }
}

export async function signOut() {
  const msal = await getMsal();
  const account = msal.getActiveAccount() || msal.getAllAccounts()[0];
  if (account) await msal.logoutPopup({ account });
  else window.location.reload();
}
