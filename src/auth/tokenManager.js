const axios = require("axios");
require("dotenv").config();
let appTokenCache = {
  token: null,
  expiresAt: 0
};

async function getAppToken() {
  if (appTokenCache.token && Date.now() < appTokenCache.expiresAt) {
    return appTokenCache.token;
  }

  const response = await axios.post(
    `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: process.env.AZURE_CLIENT_ID,
      client_secret: process.env.AZURE_CLIENT_SECRET,
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default"
    }).toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );

  appTokenCache.token = response.data.access_token;
  appTokenCache.expiresAt =
    Date.now() + (response.data.expires_in - 60) * 1000;

  return appTokenCache.token;
}

function getDelegatedToken(req) {
  const auth = req.headers.authorization;
  if (auth?.startsWith("Bearer ")) {
    return auth.slice(7);
  }
  return null;
}

module.exports = {
  getAppToken,
  getDelegatedToken
};
