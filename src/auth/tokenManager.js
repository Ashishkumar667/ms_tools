// const axios = require("axios");
// require("dotenv").config();
// let appTokenCache = {
//   token: null,
//   expiresAt: 0
// };

// async function getAppToken() {
//   if (appTokenCache.token && Date.now() < appTokenCache.expiresAt) {
//     return appTokenCache.token;
//   }

//   const response = await axios.post(
//     `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
//     new URLSearchParams({
//       client_id: process.env.AZURE_CLIENT_ID,
//       client_secret: process.env.AZURE_CLIENT_SECRET,
//       grant_type: "client_credentials",
//       scope: "https://graph.microsoft.com/.default"
//     }).toString(),
//     { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
//   );

//   appTokenCache.token = response.data.access_token;
//   appTokenCache.expiresAt =
//     Date.now() + (response.data.expires_in - 60) * 1000;

//   return appTokenCache.token;
// }

// function getDelegatedToken(req) {
//   const auth = req.headers.authorization;
//   if (auth?.startsWith("Bearer ")) {
//     return auth.slice(7);
//   }
//   return null;
// }

// module.exports = {
//   getAppToken,
//   getDelegatedToken
// };
const axios = require("axios");
const fs = require("fs");
const path = require("path");
require("dotenv").config();

let appTokenCache = {
  token: null,
  expiresAt: 0
};

// File to store user tokens persistently
const TOKEN_CACHE_FILE = path.join(__dirname, "../../tokenCache.json");

// Load cached tokens from file
function loadTokenCache() {
  try {
    if (fs.existsSync(TOKEN_CACHE_FILE)) {
      const data = fs.readFileSync(TOKEN_CACHE_FILE, "utf8");
      const cache = JSON.parse(data);
      // Convert expiresAt back to numbers
      for (const userId in cache) {
        if (cache[userId].expiresAt) {
          cache[userId].expiresAt = new Date(cache[userId].expiresAt).getTime();
        }
      }
      return new Map(Object.entries(cache));
    }
  } catch (error) {
    console.error("Failed to load token cache:", error);
  }
  return new Map();
}

// Save cached tokens to file
function saveTokenCache(cache) {
  try {
    const cacheObj = {};
    for (const [userId, data] of cache.entries()) {
      cacheObj[userId] = {
        ...data,
        expiresAt: new Date(data.expiresAt).toISOString()
      };
    }
    fs.writeFileSync(TOKEN_CACHE_FILE, JSON.stringify(cacheObj, null, 2));
  } catch (error) {
    console.error("Failed to save token cache:", error);
  }
}

// Store multiple user tokens by a user identifier
let userTokenCache = loadTokenCache();

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

/**
 * Decodes JWT to extract user ID for caching
 */
function getUserIdFromToken(token) {
  try {
    const payload = token.split('.')[1];
    const decoded = JSON.parse(Buffer.from(payload, 'base64').toString());
    return decoded.oid || decoded.sub || 'default-user';
  } catch (error) {
    console.error('Failed to decode token:', error);
    return 'default-user';
  }
}

/**
 * Refreshes an access token using refresh token
 */
async function refreshAccessToken(refreshToken) {
  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/common/oauth2/v2.0/token`,  //${process.env.AZURE_TENANT_ID || 'common'}
      new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope: [
          "openid",
        "profile",
        "offline_access",
        "User.Read",
        "User.Read.All",
        "OnlineMeetings.ReadWrite",
        "Calendars.ReadWrite",
        "TeamMember.Read.All",
        "TeamMember.ReadWrite.All",
        "Team.ReadBasic.All",
        "Chat.ReadWrite",
        "Chat.Create",
        "Chat.Read",
        // "Chat.Read.All",
        "Subscription.Read.All",
        "ChannelMessage.Send"
        ].join(" ")
      }).toString(),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    return {
      accessToken: response.data.access_token,
      refreshToken: response.data.refresh_token || refreshToken,
      expiresIn: response.data.expires_in
    };
  } catch (error) {
    console.error("Token refresh failed:", error.response?.data || error.message);
    throw new Error("Failed to refresh access token");
  }
}

/**
 * Checks if a JWT token is expired
 */
function isTokenExpired(token) {
  try {
    const payload = token.split('.')[1];
    const decoded = JSON.parse(Buffer.from(payload, 'base64').toString());
    const exp = decoded.exp * 1000; // Convert to milliseconds
    const now = Date.now();
    const isExpired = now >= exp;
    
    if (isExpired) {
      console.log(`⏱️ Token expired at ${new Date(exp).toISOString()}, now is ${new Date(now).toISOString()}`);
    }
    
    return isExpired;
  } catch (error) {
    console.error('Failed to check token expiry:', error);
    return false; // If we can't decode, assume it's valid and let Graph API reject it
  }
}

/**
 * Gets or refreshes delegated token
 */
async function getDelegatedToken(req) {
  // First, try to get token from header
  let accessToken = null;
  let refreshToken = null;
  
  const authHeader = req.headers.authorization;
  if (authHeader) {
    if (authHeader.startsWith("Bearer ")) {
      accessToken = authHeader.substring(7);
    } else {
      // Assume it's the token directly
      accessToken = authHeader;
    }
  }
  
  // Check for refresh token in custom header
  refreshToken = req.headers['x-refresh-token'] || req.body?.refreshToken;
  
  if (!accessToken) {
    // Try from body
    accessToken = req.body?.accessToken;
  }

  if (!accessToken) {
    console.log('❌ No access token provided');
    return null;
  }

  // Get user ID for cache lookup
  const userId = getUserIdFromToken(accessToken);
  console.log(`\n🔐 Processing token for user: ${userId}`);
  
  // Check if provided access token is expired
  const providedTokenExpired = isTokenExpired(accessToken);
  
  // If provided token is expired and we have a refresh token, refresh immediately
  if (providedTokenExpired && refreshToken) {
    console.log('⚠️ Provided access token is expired, refreshing with provided refresh token...');
    try {
      const refreshedTokens = await refreshAccessToken(refreshToken);
      
      // Update cache with fresh tokens
      userTokenCache.set(userId, {
        accessToken: refreshedTokens.accessToken,
        refreshToken: refreshedTokens.refreshToken,
        expiresAt: Date.now() + (refreshedTokens.expiresIn - 60) * 1000
      });
      saveTokenCache(userTokenCache);
      
      console.log(`✅ Token refreshed successfully for user: ${userId}`);
      return refreshedTokens.accessToken;
    } catch (error) {
      console.error(`❌ Failed to refresh with provided token for user ${userId}:`, error.message);
      throw error; // Throw error so user knows refresh failed
    }
  }
  
  // Check if we have cached data for this user
  const cachedData = userTokenCache.get(userId);
  
  if (cachedData) {
    console.log(`📦 Found cached data for user ${userId}, expires at: ${new Date(cachedData.expiresAt).toISOString()}`);
    
    // If cached token is still valid, return it
    if (Date.now() < cachedData.expiresAt) {
      console.log(`✅ Using cached token (still valid)`);
      return cachedData.accessToken;
    }
    
    // If expired, try to refresh using provided refresh token first, then cached
    const tokenToUse = refreshToken || cachedData.refreshToken;
    console.log(`⚠️ Cached token expired, refreshing with ${refreshToken ? 'provided' : 'cached'} refresh token...`);
    
    if (tokenToUse) {
      try {
        const refreshedTokens = await refreshAccessToken(tokenToUse);
        
        // Update cache
        userTokenCache.set(userId, {
          accessToken: refreshedTokens.accessToken,
          refreshToken: refreshedTokens.refreshToken,
          expiresAt: Date.now() + (refreshedTokens.expiresIn - 60) * 1000
        });
        saveTokenCache(userTokenCache);
        
        console.log(`✅ Token refreshed for user: ${userId}`);
        return refreshedTokens.accessToken;
      } catch (error) {
        console.error(`❌ Failed to refresh token for user ${userId}:`, error.message);
        // Remove invalid cache entry
        userTokenCache.delete(userId);
        saveTokenCache(userTokenCache);
        throw error; // Throw error instead of falling through
      }
    }
  }
  
  // Cache the new token if refresh token provided
  if (refreshToken) {
    // Check if the access token is valid before caching
    if (!providedTokenExpired) {
      userTokenCache.set(userId, {
        accessToken: accessToken,
        refreshToken: refreshToken,
        // Set expiry to 55 minutes (tokens usually valid for 1 hour)
        expiresAt: Date.now() + 55 * 60 * 1000
      });
      saveTokenCache(userTokenCache);
      console.log(`✅ Fresh tokens cached for user: ${userId}`);
      return accessToken;
    } else {
      console.log('⚠️ Not caching expired access token');
      // Try to refresh with the provided refresh token
      try {
        const refreshedTokens = await refreshAccessToken(refreshToken);
        userTokenCache.set(userId, {
          accessToken: refreshedTokens.accessToken,
          refreshToken: refreshedTokens.refreshToken,
          expiresAt: Date.now() + (refreshedTokens.expiresIn - 60) * 1000
        });
        saveTokenCache(userTokenCache);
        console.log(`✅ Refreshed and cached new tokens for user: ${userId}`);
        return refreshedTokens.accessToken;
      } catch (error) {
        console.error(`❌ Failed to refresh token:`, error.message);
        throw error;
      }
    }
  }
  
  // If we get here, return the provided token (even if expired, let Graph API reject it)
  console.log(`⚠️ Returning provided access token (no refresh available)`);
  return accessToken;
}

/**
 * Clears cached token for a user (useful for logout)
 */
function clearUserToken(token) {
  const userId = getUserIdFromToken(token);
  userTokenCache.delete(userId);
  saveTokenCache(userTokenCache);
}

module.exports = {
  getAppToken,
  getDelegatedToken,
  refreshAccessToken,
  clearUserToken
};