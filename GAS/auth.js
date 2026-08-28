const AUTH0_DOMAIN    = "vcumentors.us.auth0.com";   // from Auth0 dashboard
const AUTH0_CLIENT_ID = "JbZvr8oldo60N3zPyEfp2Pqyqn2TjzwR";              // from Auth0 dashboard
const GAS_URL = "https://script.google.com/macros/s/AKfycbw-JPye5itKs5r-qdx8kOw8ZnZ0YwcOulc1UnztyPk4h_VLWJ7pcfy8lpLjaziZMXzHFA/exec"; // your GAS URL

let auth0Client = null;

const initAuth0 = async () => {
  auth0Client = await auth0.createAuth0Client({
    domain: AUTH0_DOMAIN,
    clientId: AUTH0_CLIENT_ID,
    authorizationParams: {
      redirect_uri: window.location.origin,
      //hd: "vcu.edu"        // ← pre-filters Google to VCU accounts only
    }
  });

  // After Google login redirect, process the callback
  if (location.search.includes("code=") && location.search.includes("state=")) {
    await auth0Client.handleRedirectCallback();
    window.history.replaceState({}, document.title, window.location.pathname);
  }

  const isAuthenticated = await auth0Client.isAuthenticated();
  if (isAuthenticated) {
    const user = await auth0Client.getUser();
    onLoginSuccess(user);
  } else {
    onLoggedOut();
  }
};

const onLoginSuccess = (user) => {
  currentUserEmail = user.email;  // sets the global var used in index.html
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("mainApp").classList.remove("hidden");
  document.getElementById("userEmail").textContent = "✅ " + user.email;
  if (typeof checkUserBooking === "function") checkUserBooking();
};

const onLoggedOut = () => {
  document.getElementById("login-screen").style.display = "";
  document.getElementById("mainApp").classList.add("hidden");
};

const login = () => {
  auth0Client.loginWithRedirect({
    authorizationParams: {
      connection: "google-oauth2",
      hd: "vcu.edu"
    }
  });
};

const logout = () => {
  auth0Client.logout({
    logoutParams: { returnTo: window.location.origin }
  });
};

// This replaces ALL google.script.run calls
const callGAS = async (params = {}) => {
  const url = new URL(GAS_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.append(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error("GAS request failed: " + res.status);
  return res.json();
};

document.addEventListener("DOMContentLoaded", initAuth0);