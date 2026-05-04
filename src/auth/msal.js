// Microsoft Entra ID (MSAL) Authentication
// Lighthouse enabled - single Avontus tenant token covers all 4 tenants

const MSAL_CONFIG = {
  auth: {
    clientId: "3977e66a-cdf1-419d-9d0d-70e8cf3a76ed",
    authority: "https://login.microsoftonline.com/bd98204b-b981-4d03-8796-356d537927eb",
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

const LOGIN_SCOPES = {
  scopes: ["User.Read", "openid", "profile", "email", "https://management.azure.com/user_impersonation"]
};

const MGMT_SCOPES = {
  scopes: ["https://management.azure.com/user_impersonation"]
};

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

// Single token for all tenants - Lighthouse handles cross-tenant access
export async function getUserMgmtToken() {
  try {
    const msal = await getMsal();
    const account = msal.getActiveAccount() || msal.getAllAccounts()[0];
    if (!account) return null;
    try {
      const result = await msal.acquireTokenSilent({ ...MGMT_SCOPES, account });
      return result.accessToken;
    } catch {
      const result = await msal.acquireTokenPopup(MGMT_SCOPES);
      return result.accessToken;
    }
  } catch { return null; }
}

// Alias for backward compat
export const getAllTenantTokens = async () => {
  const token = await getUserMgmtToken();
  return token ? {
    "bd98204b-b981-4d03-8796-356d537927eb": token,
    "afafd9ca-9af6-4f95-8032-f71fc87ef9e5": token,
    "d971099d-75b2-4a01-8d0d-507161733ea5": token,
    "5e9927b8-90dd-40c9-bdb8-3283e73304c6": token,
  } : {};
};
