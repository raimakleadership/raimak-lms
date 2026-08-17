// ============================================================
//  Raimak LMS — Authentication (MSAL)
// ============================================================
const Auth = (() => {
  let msalInstance = null;
  let currentAccount = null;

  // ── Init MSAL ──────────────────────────────────────────────
  async function init() {
    try {
      msalInstance = new msal.PublicClientApplication({
        auth: {
          clientId: Config.azure.clientId,
          authority: `https://login.microsoftonline.com/${Config.azure.tenantId}`,
          redirectUri: Config.azure.redirectUri,
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: true,
        },
      });

      // 🚀 SAFETY: Future-proof for MSAL v3.x
      if (typeof msalInstance.initialize === "function") {
        await msalInstance.initialize();
      }

      // 🚀 SAFETY: Catch redirect promise errors so the app doesn't crash
      const result = await msalInstance.handleRedirectPromise();

      // Set active account from redirect result for new users
      if (result && result.account) {
        currentAccount = result.account;
        msalInstance.setActiveAccount(result.account);
      } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length) {
          currentAccount = accounts[0];
          msalInstance.setActiveAccount(accounts[0]);
        }
      }

      return result;
    } catch (err) {
      console.error("MSAL Initialization / Redirect Error:", err);

      // 🚀 SAFETY: If MSAL crashes on return, wipe the corrupted token from the URL to stop the infinite loop
      if (
        window.location.hash.includes("code=") ||
        window.location.hash.includes("error=")
      ) {
        window.history.replaceState(null, "", window.location.pathname);
      }

      return null;
    }
  }

  // ── Sign In ────────────────────────────────────────────────
  async function signIn() {
    try {
      await msalInstance.loginRedirect({ scopes: Config.scopes });
    } catch (err) {
      console.error("Sign-in error:", err);
      if (typeof UI !== "undefined" && UI.showToast) {
        UI.showToast("Sign-in failed. Please try again.", "error");
      }
    }
  }

  // ── Sign Out ───────────────────────────────────────────────
  function signOut() {
    msalInstance.logoutRedirect({
      postLogoutRedirectUri: Config.azure.redirectUri,
    });
  }

  // ── Get Access Token ───────────────────────────────────────
  async function getToken() {
    const account =
      msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if (!account) return null;

    currentAccount = account;

    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes: Config.scopes,
        account: currentAccount,
      });

      // Token came back but is empty — fallback to standard redirect
      if (!result.accessToken || result.accessToken.trim() === "") {
        console.warn("Silent token empty. Redirecting for fresh token...");
        await msalInstance.acquireTokenRedirect({ scopes: Config.scopes });
        return null;
      }

      return result.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        console.warn("Interaction required for token. Redirecting...");
        // 🚀 SAFETY: Removed `prompt: "consent"` to prevent forcing the permission screen every hour
        await msalInstance.acquireTokenRedirect({ scopes: Config.scopes });
        return null;
      }
      console.error("Critical token acquisition failure:", err);
      throw err;
    }
  }

  // ── Current User ───────────────────────────────────────────
  function getUser() {
    const account =
      msalInstance?.getActiveAccount() || msalInstance?.getAllAccounts()?.[0];
    if (!account) return null;
    currentAccount = account;
    return {
      name: account.name || account.username,
      email: account.username,
    };
  }

  function isSignedIn() {
    return !!(
      msalInstance?.getActiveAccount() || msalInstance?.getAllAccounts()?.length
    );
  }

  return { init, signIn, signOut, getToken, getUser, isSignedIn };
})();
