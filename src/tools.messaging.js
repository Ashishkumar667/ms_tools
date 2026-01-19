const { GraphClient } = require("./graphClient");

/**
 * Messaging-related tools for Microsoft Teams, backed by Microsoft Graph.
 *
 * These functions assume you already have a valid Graph access token.
 * Each one creates a GraphClient and calls the relevant Graph endpoint.
 */

/**
 * Tool: "Post an announcement to a channel"
 * API: POST /teams/{team-id}/channels/{channel-id}/messages
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.channelId
 * @param {string} params.content - HTML content of the message (body.content).
 * @returns {Promise<object>} chatMessage
 */
async function postAnnouncementToChannel(accessToken, { teamId, channelId, content }) {
  const client = new GraphClient(accessToken);

  const body = {
    body: {
      contentType: "html",
      content
    }
  };

  return client.post(`/teams/${teamId}/channels/${channelId}/messages`, body);
}

/**
 * Tool: "Send a DM in an existing chat"
 * API: POST /chats/{chat-id}/messages
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.chatId
 * @param {string} params.content - HTML or text content for the message.
 * @param {"text"|"html"} [params.contentType="html"]
 * @returns {Promise<object>} chatMessage
 */
async function sendMessageToExistingChat(accessToken, { chatId, content, contentType = "html" }) {
  const client = new GraphClient(accessToken);

  const body = {
    body: {
      contentType,
      content
    }
  };

  return client.post(`/chats/${chatId}/messages`, body);
}

/**
 * Tool: "Create a new 1:1 or group chat, then send a message"
 * APIs:
 *   - POST /chats
 *   - POST /chats/{chat-id}/messages
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {("oneOnOne"|"group")} params.chatType
 * @param {Array<{id: string}>} params.members - User IDs (AAD object IDs).
 * @param {string} params.content - Initial message content.
 * @param {"text"|"html"} [params.contentType="html"]
 * @returns {Promise<{ chat: object, message: object }>}
 */

async function getUserIdByEmail(accessToken, email) {
  const client = new GraphClient(accessToken);

  const user = await client.get(`/users/${encodeURIComponent(email)}`);
  return user.id;
}


async function createChatAndSendMessage(
  accessToken,
  { chatType, members, content, contentType = "html" }
) {
  const client = new GraphClient(accessToken);

  const resolvedMembers = await Promise.all(
    members.map(async (m) => {
      if (m.id) return m;
      if (m.email || m.userEmail) {
        const email = m.email || m.userEmail;
        const id = await getUserIdByEmail(accessToken, email);
        return { id };
      }
      throw new Error("Each member must have id or email");
    })
  );

  const chatBody = {
    chatType,
    members: resolvedMembers.map((m) => ({
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      roles: ["owner"],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.id}')`
    }))
  };
// console.log(chatBody);
  const chat = await client.post("/chats", chatBody);

  const messageBody = {
    body: {
      contentType,
      content
    }
  };

  const message = await client.post(`/chats/${chat.id}/messages`, messageBody);

  return { chat, message };
}

/**
 * Tool: "Reply to a channel message (creates thread reply)"
 * API: POST /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.channelId
 * @param {string} params.messageId - The message to reply to
 * @param {string} params.content - HTML content of the reply
 * @param {"text"|"html"} [params.contentType="html"]
 * @returns {Promise<object>} chatMessage
 */
async function replyToChannelMessage(accessToken, { teamId, channelId, messageId, content, contentType = "html" }) {
  const client = new GraphClient(accessToken);

  const body = {
    body: {
      contentType,
      content
    }
  };

  return client.post(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`, body);
}

/**
 * Tool: "Reply to a chat message"
 * API: POST /chats/{chat-id}/messages/{message-id}/replies
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.chatId
 * @param {string} params.messageId - The message to reply to
 * @param {string} params.content - HTML or text content of the reply
 * @param {"text"|"html"} [params.contentType="html"]
 * @returns {Promise<object>} chatMessage
 */
async function replyToChatMessage(accessToken, { chatId, messageId, content, contentType = "html" }) {
  const client = new GraphClient(accessToken);

  const body = {
    body: {
      contentType,
      content
    }
  };

  return client.post(`/chats/${chatId}/messages/${messageId}/replies`, body);
}
async function createChatSubscription(
  accessToken,
  { resource, notificationUrl, expirationDateTime, clientState, changeType = "created,updated", lifecycleNotificationUrl }
) {
  const client = new GraphClient(accessToken);

  const body = {
    changeType,
    notificationUrl,
    resource,
    expirationDateTime
  };

  if (clientState) {
    body.clientState = clientState;
  }

  if (lifecycleNotificationUrl) {
    body.lifecycleNotificationUrl = lifecycleNotificationUrl;
  }

  return client.post("/subscriptions", body);
}

module.exports = {
  postAnnouncementToChannel,
  sendMessageToExistingChat,
  createChatAndSendMessage,
  replyToChannelMessage,
  replyToChatMessage,
  createChatSubscription
};


