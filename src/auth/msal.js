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

const MGMT_SCOPES = { scopes: ["https://management.azure.com/user_impersonation"] };
const LOGIN_SCOPES = { scopes: ["User.Read", "openid", "profile", "email", "https://management.azure.com/user_impersonation"] };

// All tenant IDs to try acquiring tokens for
const ALL_TENANTS = [
  "bd98204b-b981-4d03-8796-356d537927eb", // Avontus Software
  "afafd9ca-9af6-4f95-8032-f71fc87ef9e5", // Places2Swim
  "d971099d-75b2-4a01-8d0d-507161733ea5", // Azure-Internal
  "5e9927b8-90dd-40c9-bdb8-3283e73304c6", // SmallBiz
];

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

// Get user's delegated token for a specific tenant
// Falls back gracefully if user doesn't have access to that tenant
export async function getUserTokenForTenant(tenantId) {
  try {
    const msal = await getMsal();
    const account = msal.getActiveAccount() || msal.getAllAccounts()[0];
    if (!account) return null;

    // Try silent first
    try {
      const result = await msal.acquireTokenSilent({
        ...MGMT_SCOPES,
        account,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      });
      return result.accessToken;
    } catch {
      // Silent failed - try popup for this specific tenant
      try {
        const result = await msal.acquireTokenPopup({
          ...MGMT_SCOPES,
          authority: `https://login.microsoftonline.com/${tenantId}`,
          prompt: "none", // Don't show UI if possible
        });
        return result.accessToken;
      } catch { return null; }
    }
  } catch { return null; }
}

// Get token for the primary (Avontus) tenant - used for most API calls
export async function getUserMgmtToken() {
  return getUserTokenForTenant("bd98204b-b981-4d03-8796-356d537927eb");
}

// Get tokens for all tenants - used for invoice fetching
export async function getAllTenantTokens() {
  const tokens = {};
  for (const tenantId of ALL_TENANTS) {
    const token = await getUserTokenForTenant(tenantId);
    if (token) tokens[tenantId] = token;
  }
  return tokens;
}
