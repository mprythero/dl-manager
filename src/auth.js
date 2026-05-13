import { PublicClientApplication, LogLevel } from "@azure/msal-browser";

// ─── REPLACE THESE TWO VALUES ─────────────────────────────────────────────────
export const CLIENT_ID  = "d0e8fd20-1463-4944-8129-48060485cef7";
export const TENANT_ID  = "a45f35e7-04f4-43d7-8feb-6b9c7174a0d0";
// ─────────────────────────────────────────────────────────────────────────────

export const msalConfig = {
  auth: {
    clientId:    CLIENT_ID,
    authority:   `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const graphScopes = [
  "Group.ReadWrite.All",
  "User.ReadBasic.All",
  "Directory.AccessAsUser.All",
];

export const msalInstance = new PublicClientApplication(msalConfig);

// Acquire token silently, fall back to popup
export async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error("Not signed in");
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes:  graphScopes,
      account: accounts[0],
    });
    return result.accessToken;
  } catch {
    const result = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return result.accessToken;
  }
}

// Graph API helper
export async function graphRequest(method, path, body) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 204) return null;
  const json = await res.json();
  if (!res.ok) throw new Error(json.error?.message || `Graph error ${res.status}`);
  return json;
}
