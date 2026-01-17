const { GraphClient } = require("./graphClient");

/**
 * Team / channel provisioning tools (create team, channels, add members, tags).
 */

/**
 * Tool: "Create a new Team from a template"
 * API: POST /teams
 *
 * This is a simplified wrapper. For production, follow the full guidance in
 * "create teams and manage members" (backing group first, then team).
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.displayName
 * @param {string} params.description
 * @param {string} params.templateId - e.g. "standard" or a custom template ID.
 * @returns {Promise<object>} Operation or created team (depending on Graph behavior).
 */
async function createTeamFromTemplate(
  accessToken,
  { displayName, description, templateId = "standard" }
) {
  const client = new GraphClient(accessToken);

  const body = {
    "template@odata.bind": `https://graph.microsoft.com/v1.0/teamsTemplates('${templateId}')`,
    displayName,
    description
  };

  return client.post("/teams", body);
}

/**
 * Tool: "Create a channel in a team"
 * API: POST /teams/{team-id}/channels
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.displayName
 * @param {string} [params.description]
 * @param {("standard"|"private"|"shared")} [params.membershipType="standard"]
 * @returns {Promise<object>} channel
 */
async function createChannelInTeam(
  accessToken,
  { teamId, displayName, description, membershipType = "standard" }
) {
  const client = new GraphClient(accessToken);

  const body = {
    displayName,
    description,
    membershipType
  };

  return client.post(`/teams/${teamId}/channels`, body);
}

/**
 * Tool: "Add a member (or owner) to a team"
 * API: POST /teams/{team-id}/members
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.userId - AAD object ID of the user.
 * @param {boolean} [params.isOwner=false]
 * @returns {Promise<object>} conversationMember
 */
async function addMemberToTeam(accessToken, { teamId, userId, isOwner = false }) {
  const client = new GraphClient(accessToken);

  const body = {
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: isOwner ? ["owner"] : ["member"],
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
  };

  return client.post(`/teams/${teamId}/members`, body);
}

/**
 * Tool: "Add a member to a private/shared channel"
 * API: POST /teams/{team-id}/channels/{channel-id}/members
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.channelId
 * @param {string} params.userId
 * @param {boolean} [params.isOwner=false]
 * @returns {Promise<object>} conversationMember
 */
async function addMemberToPrivateOrSharedChannel(
  accessToken,
  { teamId, channelId, userId, isOwner = false }
) {
  const client = new GraphClient(accessToken);

  const body = {
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: isOwner ? ["owner"] : ["member"],
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
  };

  return client.post(`/teams/${teamId}/channels/${channelId}/members`, body);
}

/**
 * Tool: "Create Teams Tags"
 * API: POST /teams/{team-id}/tags
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.teamId
 * @param {string} params.displayName
 * @param {string} [params.description]
 * @param {Array<{userId: string}>} [params.members]
 * @returns {Promise<object>} teamworkTag
 */
async function createTeamTag(
  accessToken,
  { teamId, displayName, description, members = [] }
) {
  const client = new GraphClient(accessToken);

  const body = {
    displayName,
    description,
    members: members.map((m) => ({
      userId: m.userId
    }))
  };

  return client.post(`/teams/${teamId}/tags`, body);
}

/**
 * Tool: "List Teams Tags"
 * API: GET /teams/{team-id}/tags
 *
 * @param {string} accessToken
 * @param {string} teamId
 * @returns {Promise<object>} teamworkTag collection
 */
async function listTeamTags(accessToken, teamId) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}/tags`);
}

module.exports = {
  createTeamFromTemplate,
  createChannelInTeam,
  addMemberToTeam,
  addMemberToPrivateOrSharedChannel,
  createTeamTag,
  listTeamTags
};


