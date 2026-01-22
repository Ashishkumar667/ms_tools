// const { getAppToken, getDelegatedToken } = require("./tokenManager");

// const delegatedOnlyRoutes = [
//   "/api/discovery",
//   "/api/messaging",
//   "/api/meetings"
// ];

// async function resolveGraphToken(req, res, next) {
//   try {
//     const delegatedToken = getDelegatedToken(req);

//     const requiresDelegated = delegatedOnlyRoutes.some(route =>
//       req.path.startsWith(route)
//     );

//     if (requiresDelegated) {
//       if (!delegatedToken) {
//         return res.status(401).json({
//           error: "Delegated user token required. Login first."
//         });
//       }
//       req.graphToken = delegatedToken;
//     } else {
//       // App-only token (team provisioning, admin ops)
//       req.graphToken = await getAppToken();
//     }

//     next();
//   } catch (err) {
//     res.status(500).json({
//       error: "Failed to resolve Graph token",
//       details: err.message
//     });
//   }
// }

// module.exports = resolveGraphToken;

const { getAppToken, getDelegatedToken } = require("./tokenManager");

const delegatedOnlyRoutes = [
  "/api/discovery",
  "/api/messaging",
  "/api/meetings"
];

const appOnlyRoutes = [
  "/api/teams-provisioning",
  "/api/admin"
];

/**
 * Middleware to resolve and set the appropriate Graph API token
 * Automatically handles token refresh for delegated tokens
 */
async function resolveGraphToken(req, res, next) {
  try {
    console.log(`\n🔍 Resolving token for: ${req.method} ${req.path}`);
    
    // Determine if this route requires delegated token
    const requiresDelegated = delegatedOnlyRoutes.some(route =>
      req.path.startsWith(route)
    );

    // Determine if this route requires app-only token
    const requiresAppOnly = appOnlyRoutes.some(route =>
      req.path.startsWith(route)
    );

    if (requiresDelegated) {
      // Try to get delegated token (will auto-refresh if needed)
      const delegatedToken = await getDelegatedToken(req);

      if (!delegatedToken) {
        console.log('❌ No delegated token available');
        return res.status(401).json({
          error: "Delegated user token required",
          message: "Please provide access token via Authorization header (Bearer token). Include X-Refresh-Token header or refreshToken in body for auto-refresh capability.",
          hint: "First time? Include both access_token and refresh_token. Subsequent calls will auto-refresh."
        });
      }

      req.graphToken = delegatedToken;
      console.log(`✅ Using delegated token for: ${req.path} (length: ${delegatedToken.length})`);
      
    } else if (requiresAppOnly) {
      // Use app-only token for admin operations
      req.graphToken = await getAppToken();
      console.log(`✅ Using app-only token for: ${req.path}`);
      
    } else {
      // For routes that can use either, try delegated first, then fall back to app
      const delegatedToken = await getDelegatedToken(req);
      
      if (delegatedToken) {
        req.graphToken = delegatedToken;
        console.log(`✅ Using delegated token for: ${req.path}`);
      } else {
        req.graphToken = await getAppToken();
        console.log(`✅ Using app-only token for: ${req.path}`);
      }
    }

    next();
  } catch (err) {
    console.error("❌ Token resolution error:", err);
    
    // Provide helpful error messages
    if (err.message.includes("Failed to refresh access token")) {
      return res.status(401).json({
        error: "Token refresh failed",
        message: "Your refresh token is invalid or expired. Please re-authenticate.",
        hint: "Use /auth/login to get new tokens"
      });
    }

    res.status(500).json({
      error: "Failed to resolve Graph token",
      message: err.message,
      details: err.stack
    });
  }
}

module.exports = resolveGraphToken;