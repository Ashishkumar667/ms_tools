const { GraphClient } = require("./graphClient");

/**
 * Discovery/helper tools to fetch IDs needed for other operations.
 * These endpoints help users find teamId, channelId, userId, meetingId, etc.
 */

/**
 * Get current user info
 * API: GET /me
 */
async function getCurrentUser(accessToken) {
  const client = new GraphClient(accessToken);
  return client.get("/me");
}

/**
 * List all teams the user is a member of
 * API: GET /me/joinedTeams
 */
async function listMyTeams(accessToken) {
  const client = new GraphClient(accessToken);
  return client.get("/me/joinedTeams");
}

/**
 * List all teams (requires admin permissions)
 * API: GET /teams
 */
async function listAllTeams(accessToken) {
  const client = new GraphClient(accessToken);
  return client.get("/teams");
}

/**
 * Get team details
 * API: GET /teams/{teamId}
 */
async function getTeam(accessToken, teamId) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}`);
}

/**
 * List channels in a team
 * API: GET /teams/{teamId}/channels
 */
async function listChannelsInTeam(accessToken, teamId) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}/channels`);
}

/**
 * Get channel details
 * API: GET /teams/{teamId}/channels/{channelId}
 */
async function getChannel(accessToken, { teamId, channelId }) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}/channels/${channelId}`);
}

/**
 * List chats the user is part of
 * API: GET /chats
 */
async function listMyChats(accessToken) {
  const client = new GraphClient(accessToken);
  return client.get("/chats");
}

/**
 * Get chat details
 * API: GET /chats/{chatId}
 */
async function getChat(accessToken, chatId) {
  const client = new GraphClient(accessToken);
  return client.get(`/chats/${chatId}`);
}

/**
 * List online meetings
 * API: GET /me/onlineMeetings
 */
async function listMyOnlineMeetings(accessToken, { filter, top } = {}) {
  const client = new GraphClient(accessToken);

  let url = "/me/onlineMeetings";
  const params = [];

  if (filter) {
    // encode only the filter value, not the whole $filter=
    params.push(`$filter=JoinWebUrl%20eq%20'${encodeURIComponent(filter)}'`);
  }

  if (top) {
    params.push(`$top=${top}`);
  }

  if (params.length > 0) {
    url += `?${params.join("&")}`;
  }
  console.log(params);
  console.log("Graph API URL:", url); // Should show /me/onlineMeetings?$filter=subject%20eq%20'Team%20Sync'
  return client.get(url);
}

/**
 * Get online meeting details
 * API: GET /me/onlineMeetings/{meetingId}
 */
async function getOnlineMeeting(accessToken, meetingId) {
  const client = new GraphClient(accessToken);
  return client.get(`/me/onlineMeetings/${meetingId}`);
}

/**
 * List users (requires appropriate permissions)
 * API: GET /users
 */
async function listUsers(accessToken, { filter, search, top } = {}) {
  const client = new GraphClient(accessToken);
  let url = "/users";
  const params = [];
  if (filter) params.push(`$filter=${encodeURIComponent(filter)}`);
  if (search) params.push(`$search="${encodeURIComponent(search)}"`);
  if (top) params.push(`$top=${top}`);
  if (params.length > 0) url += `?${params.join("&")}`;
  return client.get(url);
}

/**
 * Get user by ID
 * API: GET /users/{userId}
 */
async function getUser(accessToken, userId) {
  const client = new GraphClient(accessToken);
  return client.get(`/users/${userId}`);
}

/**
 * Search users by email or display name
 * API: GET /users?$filter=startswith(mail,'email') or startswith(displayName,'name')
 */
async function searchUsers(accessToken, { email, displayName } = {}) {
  const client = new GraphClient(accessToken);
  let url = "/users";
  const filters = [];
  if (email) filters.push(`startswith(mail,'${email}')`);
  if (displayName) filters.push(`startswith(displayName,'${displayName}')`);
  if (filters.length > 0) {
    url += `?$filter=${encodeURIComponent(filters.join(" or "))}`;
  }
  return client.get(url);
}

/**
 * List team members
 * API: GET /teams/{teamId}/members
 */
async function listTeamMembers(accessToken, teamId) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}/members`);
}

/**
 * List channel members (for private/shared channels)
 * API: GET /teams/{teamId}/channels/{channelId}/members
 */
async function listChannelMembers(accessToken, { teamId, channelId }) {
  const client = new GraphClient(accessToken);
  return client.get(`/teams/${teamId}/channels/${channelId}/members`);
}

module.exports = {
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
  listChannelMembers
};

