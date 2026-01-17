const { getAppToken, getDelegatedToken } = require("./tokenManager");

const delegatedOnlyRoutes = [
  "/api/discovery",
  "/api/messaging",
  "/api/meetings"
];

async function resolveGraphToken(req, res, next) {
  try {
    const delegatedToken = getDelegatedToken(req);

    const requiresDelegated = delegatedOnlyRoutes.some(route =>
      req.path.startsWith(route)
    );

    if (requiresDelegated) {
      if (!delegatedToken) {
        return res.status(401).json({
          error: "Delegated user token required. Login first."
        });
      }
      req.graphToken = delegatedToken;
    } else {
      // App-only token (team provisioning, admin ops)
      req.graphToken = await getAppToken();
    }

    next();
  } catch (err) {
    res.status(500).json({
      error: "Failed to resolve Graph token",
      details: err.message
    });
  }
}

module.exports = resolveGraphToken;
