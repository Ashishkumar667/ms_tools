require("dotenv").config();
const express = require("express");
const axios = require("axios");
const app = express();

// Middleware
app.use(express.json());

// Middleware to extract and set Graph access token
app.use((req, res, next) => {
  // Skip auth middleware for auth endpoints and root
  if (req.path.startsWith("/auth/") || req.path === "/") {
    return next();
  }

  // Extract token from Authorization header (preferred method)
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith("Bearer ")) {
    req.graphToken = authHeader.substring(7);
    return next();
  }

  // Try to get from body (for POST/PUT/PATCH requests)
  if (
    ["POST", "PUT", "PATCH"].includes(req.method) &&
    req.body &&
    req.body.accessToken
  ) {
    req.graphToken = req.body.accessToken;
    return next();
  }

  // Fallback to environment variable
  if (process.env.GRAPH_ACCESS_TOKEN) {
    req.graphToken = process.env.GRAPH_ACCESS_TOKEN;
    return next();
  }

  // If no token found and it's an API endpoint, return error
  if (req.path.startsWith("/api/")) {
    return res.status(401).json({
      error: "Missing access token",
      message:
        "Provide token via Authorization header (Bearer token), request body (accessToken for POST requests), or GRAPH_ACCESS_TOKEN env variable",
    });
  }

  // For non-API endpoints, continue without token
  next();
});

const { getAppToken } = require("./auth/tokenManager");
// Import all tools
const {
  postAnnouncementToChannel,
  sendMessageToExistingChat,
  createChatAndSendMessage,
  createChatSubscription,
} = require("./tools.messaging");
const {
  createTeamFromTemplate,
  createChannelInTeam,
  addMemberToTeam,
  addMemberToPrivateOrSharedChannel,
  createTeamTag,
  listTeamTags,
} = require("./tools.teamsProvisioning");
const {
  listMeetingTranscripts,
  getMeetingTranscript,
  getMeetingTranscriptContent,
  listMeetingRecordings,
  listMeetingAiInsights,
  getMeetingAiInsight,
} = require("./tools.meetings");
const {
  getCurrentUser,
  listMyTeams,
  listAllTeams,
  getTeam,
  listChannelsInTeam,
  getChannel,
  listMyChats,
  getChat,
  listMyOnlineMeetings,
  getOnlineMeeting,
  listUsers,
  getUser,
  searchUsers,
  listTeamMembers,
  listChannelMembers,
} = require("./tools.discovery");
const {
  resolveTeamId,
  resolveChannelId,
  resolveUserId,
  findMeetingId
} = require("./tools.idResolver");

app.get("/auth/login", (req, res) => {
  const params = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID,
    response_type: "code",
    redirect_uri: process.env.AZURE_REDIRECT_URI,
    response_mode: "query",
    scope: [
      "openid",
      "profile",
      "offline_access",

      // basic user
      "User.Read",

      // meetings
      "OnlineMeetings.Read",

      // users & org discovery
      "User.Read",

      // teams
      "TeamMember.Read.All",
      "TeamMember.ReadWrite.All",
      "Team.ReadBasic.All",

      // chat / messaging
      "Chat.ReadWrite",
      "Chat.Create",
      "ChannelMessage.Send",
    ].join(" "),
    state: "12345",
  });

  res.redirect(
    `https://login.microsoftonline.com/${
      process.env.AZURE_TENANT_ID
    }/oauth2/v2.0/authorize?${params.toString()}`
  );
});

app.get("/auth/app-token", async (req, res) => {
  const token = await getAppToken();
  res.json({
    type: "application",
    token,
  });
});

app.get("/auth/callback", async (req, res) => {
  try {
    const code = req.query.code;

    if (!code) {
      return res.status(400).json({ error: "Missing authorization code" });
    }

    const tokenRes = await axios.post(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        grant_type: "authorization_code",
        code,
        redirect_uri: process.env.AZURE_REDIRECT_URI,
      }).toString(),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );

    res.json(tokenRes.data);
  } catch (err) {
    res.status(500).json(err.response?.data || err.message);
  }
});

// Helper middleware to extract access token from Authorization header or body
function getAccessToken(req) {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith("Bearer ")) {
    return authHeader.substring(7);
  }
  return req.body.accessToken || process.env.GRAPH_ACCESS_TOKEN;
}

// ============================================================================
// DISCOVERY ENDPOINTS (to find IDs needed for other operations)
// ============================================================================

/**
 * GET /api/discovery/me
 * Get current user info
 */
app.get("/api/discovery/me", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const result = await getCurrentUser(accessToken);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams
 * List teams the user is a member of
 */
app.get("/api/discovery/teams", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { all } = req.query; // ?all=true to list all teams (requires admin)

    const result =
      all === "true"
        ? await listAllTeams(accessToken)
        : await listMyTeams(accessToken);

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams/:teamId
 * Get team details
 */
app.get("/api/discovery/teams/:teamId", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { teamId } = req.params;
    const resolvedTeamId = await resolveTeamId(accessToken, teamId);
    if (!resolvedTeamId) {
      return res.status(404).json({ error: `Team not found: ${teamId}` });
    }
    const result = await getTeam(accessToken, resolvedTeamId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams/:teamId/channels
 * List channels in a team
 */
app.get("/api/discovery/teams/:teamId/channels", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { teamId } = req.params;
    const resolvedTeamId = await resolveTeamId(accessToken, teamId);
    if (!resolvedTeamId) {
      return res.status(404).json({ error: `Team not found: ${teamId}` });
    }
    const result = await listChannelsInTeam(accessToken, resolvedTeamId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams/:teamId/channels/:channelId
 * Get channel details
 */
app.get(
  "/api/discovery/teams/:teamId/channels/:channelId",
  async (req, res) => {
    try {
      const accessToken = req.graphToken;
      const { teamId, channelId } = req.params;
      const resolvedTeamId = await resolveTeamId(accessToken, teamId);
      if (!resolvedTeamId) {
        return res.status(404).json({ error: `Team not found: ${teamId}` });
      }
      const channelInfo = await resolveChannelId(accessToken, channelId, resolvedTeamId);
      if (!channelInfo || !channelInfo.channelId) {
        return res.status(404).json({ error: `Channel not found: ${channelId}` });
      }
      const result = await getChannel(accessToken, { teamId: channelInfo.teamId, channelId: channelInfo.channelId });
      res.json({ success: true, data: result });
    } catch (error) {
      res.status(error.response?.status || 500).json({
        error: error.message,
        details: error.response?.data || error.stack,
      });
    }
  }
);

/**
 * GET /api/discovery/chats
 * List chats the user is part of
 */
app.get("/api/discovery/chats", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const result = await listMyChats(accessToken);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/chats/:chatId
 * Get chat details
 */
app.get("/api/discovery/chats/:chatId", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { chatId } = req.params;
    const result = await getChat(accessToken, chatId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/meetings
 * List online meetings
 * Query params: filter, top
 */
app.get("/api/discovery/meetings", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { filter, top } = req.query;
    const result = await listMyOnlineMeetings(accessToken, { filter, top });
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/meetings/:meetingId
 * Get online meeting details
 */
app.get("/api/discovery/meetings/:meetingId", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { meetingId } = req.params;
    const result = await getOnlineMeeting(accessToken, meetingId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/users
 * List users
 * Query params: filter, search, top
 */
app.get("/api/discovery/users", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { filter, search, top, email, displayName } = req.query;

    let result;
    if (email || displayName) {
      result = await searchUsers(accessToken, { email, displayName });
    } else {
      result = await listUsers(accessToken, { filter, search, top });
    }

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/users/:userId
 * Get user by ID
 */
app.get("/api/discovery/users/:userId", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { userId } = req.params;
    const resolvedUserId = await resolveUserId(accessToken, userId);
    if (!resolvedUserId) {
      return res.status(404).json({ error: `User not found: ${userId}` });
    }
    const result = await getUser(accessToken, resolvedUserId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams/:teamId/members
 * List team members
 */
app.get("/api/discovery/teams/:teamId/members", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { teamId } = req.params;
    const resolvedTeamId = await resolveTeamId(accessToken, teamId);
    if (!resolvedTeamId) {
      return res.status(404).json({ error: `Team not found: ${teamId}` });
    }
    const result = await listTeamMembers(accessToken, resolvedTeamId);
    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/discovery/teams/:teamId/channels/:channelId/members
 * List channel members (for private/shared channels)
 */
app.get(
  "/api/discovery/teams/:teamId/channels/:channelId/members",
  async (req, res) => {
    try {
      const accessToken = req.graphToken;
      const { teamId, channelId } = req.params;
      const resolvedTeamId = await resolveTeamId(accessToken, teamId);
      if (!resolvedTeamId) {
        return res.status(404).json({ error: `Team not found: ${teamId}` });
      }
      const channelInfo = await resolveChannelId(accessToken, channelId, resolvedTeamId);
      if (!channelInfo || !channelInfo.channelId) {
        return res.status(404).json({ error: `Channel not found: ${channelId}` });
      }
      const result = await listChannelMembers(accessToken, {
        teamId: channelInfo.teamId,
        channelId: channelInfo.channelId,
      });
      res.json({ success: true, data: result });
    } catch (error) {
      res.status(error.response?.status || 500).json({
        error: error.message,
        details: error.response?.data || error.stack,
      });
    }
  }
);

// ============================================================================
// WEBHOOK / WORKFLOWS ENDPOINTS (no Graph auth required)
// ============================================================================

/**
 * Tool H1: Inbound webhook endpoint for alerts
 *
 * External systems (CI/CD, monitoring, incident mgmt) POST here.
 * This server then forwards the alert to a Teams channel using a
 * pre-configured Teams incoming webhook URL (no Graph auth).
 *
 * Environment variables:
 *   - TEAMS_INCOMING_WEBHOOK_URL: the full Teams incoming webhook URL
 *   - INBOUND_WEBHOOK_SECRET (optional): shared secret to validate callers
 */
app.post("/api/webhooks/inbound/alert", async (req, res) => {
  try {
    const webhookUrl = process.env.TEAMS_INCOMING_WEBHOOK_URL;
    if (!webhookUrl) {
      return res.status(500).json({
        error:
          "TEAMS_INCOMING_WEBHOOK_URL is not configured in environment variables",
      });
    }

    const expectedSecret = process.env.INBOUND_WEBHOOK_SECRET;
    const providedSecret = req.headers["x-shared-secret"];
    if (expectedSecret && providedSecret !== expectedSecret) {
      return res.status(401).json({ error: "Invalid shared secret" });
    }

    const {
      title = "Alert",
      summary,
      severity = "info",
      source = "external-system",
      details,
    } = req.body || {};

    if (!summary) {
      return res.status(400).json({
        error: "Missing required field: summary",
      });
    }

    // Simple Teams MessageCard payload (for classic incoming webhook)
    const card = {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      summary,
      themeColor:
        severity === "critical"
          ? "FF0000"
          : severity === "warning"
          ? "FFA500"
          : "0076D7",
      title,
      sections: [
        {
          activityTitle: `Source: ${source}`,
          text: summary,
          facts: details
            ? Object.entries(details).map(([name, value]) => ({
                name,
                value: String(value),
              }))
            : [],
        },
      ],
    };

    await axios.post(webhookUrl, card);

    res.json({ success: true });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * Tool H2: Outgoing webhook command bridge
 *
 * Configure this URL as an Outgoing Webhook in a Team.
 * When a user types @YourWebhook some text, Teams will POST here.
 * You must respond within ~10 seconds with a simple payload.
 *
 * NOTE: For brevity we are not validating HMAC signatures here.
 */
app.post("/apii/teams/webhooks/outgoing", async (req, res) => {
  try {
    // Handle Teams webhook validation challenge
    // Teams sends this during webhook creation to verify the endpoint
    if (req.body && req.body.type === "verification") {
      // Return the challenge token to validate the webhook
      return res.status(200).send(req.body.value);
    }

    const { text, from, channelId } = req.body || {};

    const userName = from?.name || "user";
    const replyText = `Hi ${userName}, you said: "${text || ""}" in channel ${
      channelId || ""
    }.`;

    // Teams expects JSON with a "text" field (and optional attachments)
    res.json({
      text: replyText,
    });
  } catch (error) {
    console.error("Webhook error:", error);
    res.status(500).json({
      error: error.message,
      details: error.stack,
    });
  }
});

// ============================================================================
// MESSAGING ENDPOINTS
// ============================================================================

/**
 * POST /api/messaging/channel/announcement
 * Post an announcement to a channel
 */
app.post("/api/messaging/channel/announcement", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { teamId, teamName, channelId, channelName, content } = req.body;

    if (!content) {
      return res.status(400).json({
        error: "Missing required field: content",
      });
    }

    // Auto-resolve teamId from name if needed
    const resolvedTeamId = teamId || await resolveTeamId(accessToken, teamName);
    if (!resolvedTeamId) {
      return res.status(400).json({
        error: "Missing teamId or teamName. Provide either teamId (GUID) or teamName (string)",
      });
    }

    // Auto-resolve channelId from name if needed
    const channelInfo = channelId 
      ? { channelId, teamId: resolvedTeamId }
      : await resolveChannelId(accessToken, channelName, resolvedTeamId);
    
    if (!channelInfo || !channelInfo.channelId) {
      return res.status(400).json({
        error: "Missing channelId or channelName. Provide either channelId (GUID) or channelName (string)",
      });
    }

    const result = await postAnnouncementToChannel(accessToken, {
      teamId: channelInfo.teamId,
      channelId: channelInfo.channelId,
      content,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/messaging/chat/send
 * Send a DM in an existing chat
 */
app.post("/api/messaging/chat/send", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { chatId, content, contentType = "html" } = req.body;

    if (!chatId || !content) {
      return res.status(400).json({
        error: "Missing required fields: chatId, content",
      });
    }

    const result = await sendMessageToExistingChat(accessToken, {
      chatId,
      content,
      contentType,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/messaging/chat/create-and-send
 * Create a new 1:1 or group chat, then send a message
 */
app.post("/api/messaging/chat/create-and-send", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { chatType, members, content, contentType = "html" } = req.body;

    if (!chatType || !members || !Array.isArray(members) || !content) {
      return res.status(400).json({
        error: "Missing required fields: chatType, members (array), content",
      });
    }

    const result = await createChatAndSendMessage(accessToken, {
      chatType,
      members,
      content,
      contentType,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/messaging/subscriptions/create
 * Create a chat change watcher subscription
 */
app.post("/api/messaging/subscriptions/create", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const {
      resource,
      notificationUrl,
      expirationDateTime,
      clientState,
      changeType = "created,updated",
    } = req.body;

    if (!resource || !notificationUrl || !expirationDateTime) {
      return res.status(400).json({
        error:
          "Missing required fields: resource, notificationUrl, expirationDateTime",
      });
    }

    const result = await createChatSubscription(accessToken, {
      resource,
      notificationUrl,
      expirationDateTime,
      clientState,
      changeType,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

// ============================================================================
// TEAM / CHANNEL PROVISIONING ENDPOINTS
// ============================================================================

/**
 * POST /api/teams/create
 * Create a new Team from a template
 */
app.post("/api/teams/create", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { displayName, description, templateId = "standard" } = req.body;

    if (!displayName) {
      return res.status(400).json({
        error: "Missing required field: displayName",
      });
    }

    const result = await createTeamFromTemplate(accessToken, {
      displayName,
      description,
      templateId,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/teams/:teamId/channels/create
 * Create a channel in a team
 */
app.post("/api/teams/channels/create", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    // const urlTeamId = req.params.teamId;
    const { teamId, teamName, displayName, description, membershipType = "standard" } = req.body;

    if (!displayName) {
      return res.status(400).json({
        error: "Missing required field: displayName",
      });
    }

    // Auto-resolve teamId from URL param or body (body takes precedence)
    const teamIdOrName = teamId || teamName ;
    const resolvedTeamId = await resolveTeamId(accessToken, teamIdOrName);
    
    if (!resolvedTeamId) {
      return res.status(400).json({
        error: `Team not found: ${teamIdOrName}. Provide either teamId (GUID) or teamName (string)`,
      });
    }

    const result = await createChannelInTeam(accessToken, {
      teamId: resolvedTeamId,
      displayName,
      description,
      membershipType,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/teams/:teamId/members/add
 * Add a member (or owner) to a team
 */
app.post("/api/teams/members/add", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    // const urlTeamId = req.params.teamId;
    const { teamId, teamName, userId, userEmail, userName, isOwner = false } = req.body;

    // Auto-resolve teamId
    const teamIdOrName = teamId || teamName ;
    const resolvedTeamId = await resolveTeamId(accessToken, teamIdOrName);
    
    if (!resolvedTeamId) {
      return res.status(400).json({
        error: `Team not found: ${teamIdOrName}. Provide either teamId (GUID) or teamName (string)`,
      });
    }

    // Auto-resolve userId from email or name
    const userIdOrEmailOrName = userId || userEmail || userName;
    if (!userIdOrEmailOrName) {
      return res.status(400).json({
        error: "Missing required field: userId, userEmail, or userName",
      });
    }

    const resolvedUserId = await resolveUserId(accessToken, userIdOrEmailOrName);
    if (!resolvedUserId) {
      return res.status(400).json({
        error: `User not found: ${userIdOrEmailOrName}`,
      });
    }

    const result = await addMemberToTeam(accessToken, {
      teamId: resolvedTeamId,
      userId: resolvedUserId,
      isOwner,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * POST /api/teams/:teamId/channels/:channelId/members/add
 * Add a member to a private/shared channel
 */
app.post(
  "/api/teams/channels/members/add",
  async (req, res) => {
    try {
      const accessToken = req.graphToken;
      // const urlTeamId = req.params.teamId;
      // const urlChannelId = req.params.channelId;
      const { teamId, teamName, channelId, channelName, userId, userEmail, userName, isOwner = false } = req.body;

      // Auto-resolve teamId
      const teamIdOrName = teamId || teamName ;
      const resolvedTeamId = await resolveTeamId(accessToken, teamIdOrName);
      if (!resolvedTeamId) {
        return res.status(404).json({ error: `Team not found: ${teamIdOrName}` });
      }

      // Auto-resolve channelId
      const channelIdOrName = channelId || channelName ;
      const channelInfo = await resolveChannelId(accessToken, channelIdOrName, resolvedTeamId);
      if (!channelInfo || !channelInfo.channelId) {
        return res.status(404).json({ error: `Channel not found: ${channelIdOrName}` });
      }

      // Auto-resolve userId
      const userIdOrEmailOrName = userId || userEmail || userName;
      if (!userIdOrEmailOrName) {
        return res.status(400).json({
          error: "Missing required field: userId, userEmail, or userName",
        });
      }

      const resolvedUserId = await resolveUserId(accessToken, userIdOrEmailOrName);
      if (!resolvedUserId) {
        return res.status(404).json({ error: `User not found: ${userIdOrEmailOrName}` });
      }

      const result = await addMemberToPrivateOrSharedChannel(accessToken, {
        teamId: channelInfo.teamId,
        channelId: channelInfo.channelId,
        userId: resolvedUserId,
        isOwner,
      });

      res.json({ success: true, data: result });
    } catch (error) {
      res.status(error.response?.status || 500).json({
        error: error.message,
        details: error.response?.data || error.stack,
      });
    }
  }
);

/**
 * POST /api/teams/:teamId/tags/create
 * Create a Teams tag
 */
app.post("/api/teams/tags/create", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    // const urlTeamId = req.params.teamId;
    const { teamId, teamName, displayName, description, members = [] } = req.body;

    if (!displayName) {
      return res.status(400).json({
        error: "Missing required field: displayName",
      });
    }

    const teamIdOrName = teamId || teamName ;
    const resolvedTeamId = await resolveTeamId(accessToken, teamIdOrName);
    if (!resolvedTeamId) {
      return res.status(404).json({ error: `Team not found: ${teamIdOrName}` });
    }

    const result = await createTeamTag(accessToken, {
      teamId: resolvedTeamId,
      displayName,
      description,
      members,
    });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/teams/:teamId/tags
 * List Teams tags
 */
app.get("/api/teams/tags", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { teamName } = req.body;

    const resolvedTeamId = await resolveTeamId(accessToken, teamName);
    if (!resolvedTeamId) {
      return res.status(404).json({ error: `Team not found: ${teamName}` });
    }

    const result = await listTeamTags(accessToken, resolvedTeamId);

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

// ============================================================================
// MEETING ENDPOINTS
// ============================================================================

/**
 * GET /api/meetings/:meetingId/transcripts
 * List transcripts for a scheduled meeting
 */
app.get("/api/meetings/transcripts", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { meetingId, title, organizerEmail, afterDate } = req.query;

    // Auto-resolve meetingId from query params if provided
    let resolvedMeetingId = meetingId;
    if (!resolvedMeetingId && (title || organizerEmail || afterDate)) {
      resolvedMeetingId = await findMeetingId(accessToken, { 
        title, 
        organizerEmail, 
        afterDate 
      });
    }

    if (!resolvedMeetingId) {
      return res.status(400).json({
        error: "Missing meetingId or meeting search criteria (title, organizerEmail, afterDate)",
      });
    }

    const result = await listMeetingTranscripts(accessToken, { meetingId: resolvedMeetingId });

    res.json({ success: true, data: result });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/meetings/:meetingId/transcripts/:transcriptId
 * Get transcript metadata
 */

async function getTranscriptContent(accessToken, transcriptContentUrl) {
  const response = await axios.get(transcriptContentUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "text/vtt",
    },
  });
  console.log(response.data);
  return response.data;
}

app.get(
  "/api/meetings/transcripts/:transcriptId",
  async (req, res) => {
    try {
      const accessToken = req.graphToken;
      const { transcriptId } = req.params;
      const { meetingId, title, organizerEmail, afterDate } = req.query;

      // Auto-resolve meetingId from query params if provided
      let resolvedMeetingId = meetingId;
      if (!resolvedMeetingId && (title || organizerEmail || afterDate)) {
        resolvedMeetingId = await findMeetingId(accessToken, { 
          title, 
          organizerEmail, 
          afterDate 
        });
      }

      if (!resolvedMeetingId) {
        return res.status(400).json({
          error: "Missing meetingId or meeting search criteria (title, organizerEmail, afterDate) in query params",
        });
      }

      const result = await getMeetingTranscript(accessToken, {
        meetingId: resolvedMeetingId,
        transcriptId,
      });
      const content = await getTranscriptContent(
        accessToken,
        result.transcriptContentUrl
      );
      res.json({ success: true, data: result, transcriptContent: content });
    } catch (error) {
      res.status(error.response?.status || 500).json({
        error: error.message,
        details: error.response?.data || error.stack,
      });
    }
  }
);


/**
 * GET /api/meetings/:meetingId/recordings
 * List recordings for a scheduled meeting
 */

const {
  uploadVideoFromUrlToCloudinary,
  uploadVideoToCloudinary,
} = require("./utils/cloudinary");

const getRecordingContent = async (accessToken, recordingContentUrl) => {
  if (!recordingContentUrl) {
    throw new Error("recordingContentUrl is missing");
  }

  const response = await axios.get(recordingContentUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "video/mp4",
    },
    responseType: "arraybuffer",
  });

  return response.data;
};

app.get("/api/meetings/recordings", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { meetingId, title, organizerEmail, afterDate } = req.query;

    // Auto-resolve meetingId from query params if provided
    let resolvedMeetingId = meetingId;
    if (!resolvedMeetingId && (title || organizerEmail || afterDate)) {
      resolvedMeetingId = await findMeetingId(accessToken, { 
        title, 
        organizerEmail, 
        afterDate 
      });
    }

    if (!resolvedMeetingId) {
      return res.status(400).json({
        error: "Missing meetingId or meeting search criteria (title, organizerEmail, afterDate) in query params",
      });
    }

    const result = await listMeetingRecordings(accessToken, { meetingId: resolvedMeetingId });

    // 🔥 Correct access
    const recording = result?.value?.[0];

    if (!recording?.recordingContentUrl) {
      return res.status(404).json({
        error: "Recording not found",
        result,
      });
    }

    // Check if Cloudinary is configured
    if (!process.env.CLOUDINARY_CLOUD_NAME || !process.env.CLOUDINARY_API_KEY || !process.env.CLOUDINARY_API_SECRET) {
      return res.status(500).json({
        error: "Cloudinary not configured",
        message: "Please set CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, and CLOUDINARY_API_SECRET environment variables",
      });
    }

    // Teams recordings require authentication, so we need to download first, then upload
    // Download the recording to buffer
    const videoBuffer = await getRecordingContent(
      accessToken,
      recording.recordingContentUrl
    );

    // Upload to Cloudinary
    const publicId = `meeting-${resolvedMeetingId}-${Date.now()}`;
    const cloudinaryResult = await uploadVideoToCloudinary(videoBuffer, {
      publicId: publicId,
      folder: "teams-recordings",
    });

    // Return Cloudinary URL
    res.json({
      success: true,
      data: {
        meetingId: resolvedMeetingId,
        recordingId: recording.id,
        cloudinaryUrl: cloudinaryResult.secure_url,
        publicId: cloudinaryResult.public_id,
        format: cloudinaryResult.format,
        duration: cloudinaryResult.duration,
        bytes: cloudinaryResult.bytes,
        createdAt: cloudinaryResult.created_at,
        originalRecordingUrl: recording.recordingContentUrl, // Keep for reference
      },
      message: "Recording uploaded to Cloudinary successfully",
    });
  } catch (error) {
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/meetings/:meetingId/ai-insights
 * List AI insights for a meeting (filtered by type)
 * Query params: userId, aiInsightType (meetingSummary, callToAction, mention)
 */

function filterInsightsByType(insights, type) {
  const typeMap = {
    summary: "#microsoft.graph.callAiSummary",
    actionItems: "#microsoft.graph.callAiActionItem",
    decisions: "#microsoft.graph.callAiDecision",
    topics: "#microsoft.graph.callAiTopic",
    speakers: "#microsoft.graph.callAiSpeakerTimeline",
  };

  const odataType = typeMap[type];
  if (!odataType) return [];

  return insights.value.filter((i) => i["@odata.type"] === odataType);
}

app.get("/api/meetings/ai-insights", async (req, res) => {
  try {
    const accessToken = req.graphToken;
    const { meetingId, title, organizerEmail, afterDate, userId, userEmail, userName, type } = req.query;

    if (!type) {
      return res.status(400).json({
        error: "Missing required query param: type",
      });
    }

    // Auto-resolve meetingId from query params if provided
    let resolvedMeetingId = meetingId;
    if (!resolvedMeetingId && (title || organizerEmail || afterDate)) {
      resolvedMeetingId = await findMeetingId(accessToken, { 
        title, 
        organizerEmail, 
        afterDate 
      });
    }

    if (!resolvedMeetingId) {
      return res.status(400).json({
        error: "Missing meetingId or meeting search criteria (title, organizerEmail, afterDate) in query params",
      });
    }

    // Auto-resolve userId from email or name
    const userIdOrEmailOrName = userId || userEmail || userName;
    if (!userIdOrEmailOrName) {
      return res.status(400).json({
        error: "Missing required query param: userId, userEmail, or userName",
      });
    }

    const resolvedUserId = await resolveUserId(accessToken, userIdOrEmailOrName);
    if (!resolvedUserId) {
      return res.status(404).json({ error: `User not found: ${userIdOrEmailOrName}` });
    }

    const result = await listMeetingAiInsights(accessToken, {
      userId: resolvedUserId,
      meetingId: resolvedMeetingId,
    });

    // Note: Error handling moved to catch block below

    const filtered = filterInsightsByType(result, type);

    res.json({
      success: true,
      total: filtered.length,
      data: filtered,
    });
  } catch (error) {
    // Handle Copilot license error
    if (
      error.response?.status === 403 &&
      error.response?.data?.error?.message?.includes("Copilot license")
    ) {
      return res.status(403).json({
        error: "Copilot not licensed",
        message:
          "This user does not have Microsoft Copilot enabled. AI insights are unavailable.",
      });
    }
    res.status(error.response?.status || 500).json({
      error: error.message,
      details: error.response?.data || error.stack,
    });
  }
});

/**
 * GET /api/meetings/:meetingId/ai-insights/:aiInsightId
 * Get a specific AI insight
 * Query params: userId
 */
app.get(
  "/api/meetings/ai-insights/:aiInsightId",
  async (req, res) => {
    try {
      const accessToken = req.graphToken;
      const { aiInsightId } = req.params;
      const { meetingId, title, organizerEmail, afterDate, userId, userEmail, userName } = req.query;

      // Auto-resolve meetingId from query params if provided
      let resolvedMeetingId = meetingId;
      if (!resolvedMeetingId && (title || organizerEmail || afterDate)) {
        resolvedMeetingId = await findMeetingId(accessToken, { 
          title, 
          organizerEmail, 
          afterDate 
        });
      }

      if (!resolvedMeetingId) {
        return res.status(400).json({
          error: "Missing meetingId or meeting search criteria (title, organizerEmail, afterDate) in query params",
        });
      }

      // Auto-resolve userId from email or name
      const userIdOrEmailOrName = userId || userEmail || userName;
      if (!userIdOrEmailOrName) {
        return res.status(400).json({
          error: "Missing required query param: userId, userEmail, or userName",
        });
      }

      const resolvedUserId = await resolveUserId(accessToken, userIdOrEmailOrName);
      if (!resolvedUserId) {
        return res.status(404).json({ error: `User not found: ${userIdOrEmailOrName}` });
      }

      const result = await getMeetingAiInsight(accessToken, {
        userId: resolvedUserId,
        meetingId: resolvedMeetingId,
        aiInsightId,
      });

      res.json({ success: true, data: result });
    } catch (error) {
      // Handle Copilot license error
      if (
        error.response?.status === 403 &&
        error.response?.data?.error?.message?.includes("Copilot license")
      ) {
        return res.status(403).json({
          error: "Copilot not licensed",
          message:
            "This user does not have Microsoft Copilot enabled. AI insights are unavailable.",
        });
      }

      res.status(error.response?.status || 500).json({
        error: error.message,
        details: error.response?.data || error.stack,
      });
    }
  }
);

// ============================================================================
// HEALTH CHECK
// ============================================================================

app.get("/", (req, res) => {
  res.json({
    message: "Microsoft Teams Graph API Tools Server",
    version: "1.0.0",
    endpoints: {
      discovery: [
        "GET /api/discovery/me - Get current user info",
        "GET /api/discovery/teams - List teams (?all=true for all teams)",
        "GET /api/discovery/teams/:teamId - Get team details",
        "GET /api/discovery/teams/:teamId/channels - List channels in team",
        "GET /api/discovery/teams/:teamId/channels/:channelId - Get channel details",
        "GET /api/discovery/chats - List chats",
        "GET /api/discovery/chats/:chatId - Get chat details",
        "GET /api/discovery/meetings - List online meetings",
        "GET /api/discovery/meetings/:meetingId - Get meeting details",
        "GET /api/discovery/users - List/search users (?email=... or ?displayName=...)",
        "GET /api/discovery/users/:userId - Get user details",
        "GET /api/discovery/teams/:teamId/members - List team members",
        "GET /api/discovery/teams/:teamId/channels/:channelId/members - List channel members",
      ],
      webhooks: [
        "POST /api/webhooks/inbound/alert - Inbound alert webhook to Teams incoming webhook",
        "POST /api/webhooks/outgoing - Outgoing webhook command bridge for @mentions",
      ],
      messaging: [
        "POST /api/messaging/channel/announcement",
        "POST /api/messaging/chat/send",
        "POST /api/messaging/chat/create-and-send",
        "POST /api/messaging/subscriptions/create",
      ],
      teams: [
        "POST /api/teams/create",
        "POST /api/teams/:teamId/channels/create",
        "POST /api/teams/:teamId/members/add",
        "POST /api/teams/:teamId/channels/:channelId/members/add",
        "POST /api/teams/:teamId/tags/create",
        "GET /api/teams/:teamId/tags",
      ],
      meetings: [
        "GET /api/meetings/:meetingId/transcripts",
        "GET /api/meetings/:meetingId/transcripts/:transcriptId",
        "GET /api/meetings/:meetingId/transcripts/:transcriptId/content",
        "GET /api/meetings/:meetingId/recordings",
        "GET /api/meetings/:meetingId/ai-insights?userId=...&aiInsightType=...",
        "GET /api/meetings/:meetingId/ai-insights/:aiInsightId?userId=...",
      ],
    },
    usage: {
      step1: "Start with GET /api/discovery/me to get your user info",
      step2: "Use GET /api/discovery/teams to find team IDs",
      step3:
        "Use GET /api/discovery/teams/:teamId/channels to find channel IDs",
      step4: "Use GET /api/discovery/users to find user IDs",
      step5: "Use GET /api/discovery/meetings to find meeting IDs",
      step6:
        "Configure TEAMS_INCOMING_WEBHOOK_URL to use /api/webhooks/inbound/alert",
      step7:
        "Configure an Outgoing Webhook in Teams to call /api/webhooks/outgoing",
    },
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(
    `🚀 Microsoft Teams Graph API Tools Server running on http://localhost:${PORT}`
  );
  console.log(`📖 API documentation available at http://localhost:${PORT}`);
});
