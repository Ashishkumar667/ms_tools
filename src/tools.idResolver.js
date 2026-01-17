const { GraphClient } = require("./graphClient");

/**
 * ID Resolution utilities - automatically find IDs from human-readable names
 * These functions help endpoints accept names/emails/titles instead of requiring IDs upfront
 */

/**
 * Find team ID by team name (displayName)
 */
async function findTeamIdByName(accessToken, teamName) {
  const client = new GraphClient(accessToken);
  
  try {
    // Search user's joined teams
    const teams = await client.get("/me/joinedTeams");
    
    if (teams.value) {
      const team = teams.value.find(
        t => t.displayName?.toLowerCase() === teamName.toLowerCase()
      );
      if (team) {
        return team.id;
      }
    }
    
    // If not found, try all teams (requires admin)
    try {
      const allTeams = await client.get("/teams");
      if (allTeams.value) {
        const team = allTeams.value.find(
          t => t.displayName?.toLowerCase() === teamName.toLowerCase()
        );
        if (team) {
          return team.id;
        }
      }
    } catch (err) {
      // May not have permissions, ignore
    }
    
    return null;
  } catch (error) {
    throw new Error(`Failed to find team: ${error.message}`);
  }
}

/**
 * Find channel ID by channel name within a team
 * Accepts either teamId or teamName
 */
async function findChannelIdByName(accessToken, channelName, teamIdOrName) {
  const client = new GraphClient(accessToken);
  
  // If teamIdOrName looks like an ID (GUID format), use it directly
  let teamId = teamIdOrName;
  if (!teamIdOrName.match(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i)) {
    // It's a name, resolve to ID
    teamId = await findTeamIdByName(accessToken, teamIdOrName);
    if (!teamId) {
      throw new Error(`Team not found: ${teamIdOrName}`);
    }
  }
  
  try {
    const channels = await client.get(`/teams/${teamId}/channels`);
    
    if (channels.value) {
      const channel = channels.value.find(
        c => c.displayName?.toLowerCase() === channelName.toLowerCase()
      );
      if (channel) {
        return { channelId: channel.id, teamId };
      }
    }
    
    return null;
  } catch (error) {
    throw new Error(`Failed to find channel: ${error.message}`);
  }
}

/**
 * Find user ID by email or display name
 */
async function findUserId(accessToken, emailOrName) {
  const client = new GraphClient(accessToken);
  
  try {
    // If it looks like an email, search by email
    if (emailOrName.includes("@")) {
      const users = await client.get(`/users?$filter=startswith(mail,'${emailOrName}') or mail eq '${emailOrName}'`);
      if (users.value && users.value.length > 0) {
        return users.value[0].id;
      }
    }
    
    // Search by display name
    const users = await client.get(`/users?$filter=startswith(displayName,'${emailOrName}')`);
    if (users.value && users.value.length > 0) {
      return users.value[0].id;
    }
    
    // Try user principal name
    try {
      const user = await client.get(`/users/${emailOrName}`);
      return user.id;
    } catch (err) {
      // Not a UPN, continue
    }
    
    return null;
  } catch (error) {
    throw new Error(`Failed to find user: ${error.message}`);
  }
}

/**
 * Find meeting ID by meeting title, date, or organizer
 * Returns the most recent matching meeting
 */
async function findMeetingId(accessToken, { title, organizerEmail, afterDate } = {}) {
  const client = new GraphClient(accessToken);
  
  try {
    let meetings = await client.get("/me/onlineMeetings");
    
    if (!meetings.value || meetings.value.length === 0) {
      return null;
    }
    
    // Filter meetings
    let filtered = meetings.value;
    
    if (title) {
      filtered = filtered.filter(
        m => m.subject?.toLowerCase().includes(title.toLowerCase())
      );
    }
    
    if (organizerEmail) {
      const organizerId = await findUserId(accessToken, organizerEmail);
      if (organizerId) {
        filtered = filtered.filter(m => m.participants?.organizer?.identity?.user?.id === organizerId);
      }
    }
    
    if (afterDate) {
      const after = new Date(afterDate);
      filtered = filtered.filter(m => {
        const startTime = m.startDateTime ? new Date(m.startDateTime) : null;
        return startTime && startTime >= after;
      });
    }
    
    // Sort by start time (most recent first) and return first match
    filtered.sort((a, b) => {
      const aTime = a.startDateTime ? new Date(a.startDateTime) : new Date(0);
      const bTime = b.startDateTime ? new Date(b.startDateTime) : new Date(0);
      return bTime - aTime;
    });
    
    return filtered.length > 0 ? filtered[0].id : null;
  } catch (error) {
    throw new Error(`Failed to find meeting: ${error.message}`);
  }
}

/**
 * Find chat ID by participant emails or chat topic
 */
async function findChatId(accessToken, { participantEmails, topic } = {}) {
  const client = new GraphClient(accessToken);
  
  try {
    const chats = await client.get("/chats");
    
    if (!chats.value || chats.value.length === 0) {
      return null;
    }
    
    let filtered = chats.value;
    
    // Filter by topic if provided
    if (topic) {
      filtered = filtered.filter(
        c => c.topic?.toLowerCase().includes(topic.toLowerCase())
      );
    }
    
    // Filter by participants if provided
    if (participantEmails && participantEmails.length > 0) {
      const participantIds = await Promise.all(
        participantEmails.map(email => findUserId(accessToken, email))
      );
      
      filtered = filtered.filter(chat => {
        const chatMemberIds = chat.members?.map(m => m.userId || m.id) || [];
        return participantIds.every(id => id && chatMemberIds.includes(id));
      });
    }
    
    // Return first match (or most recent)
    return filtered.length > 0 ? filtered[0].id : null;
  } catch (error) {
    throw new Error(`Failed to find chat: ${error.message}`);
  }
}

/**
 * Resolve teamId - accepts either ID or name
 */
async function resolveTeamId(accessToken, teamIdOrName) {
  if (!teamIdOrName) return null;
  
  // If it's a GUID, assume it's already an ID
  if (teamIdOrName.match(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i)) {
    return teamIdOrName;
  }
  
  // Otherwise, treat as name and resolve
  return await findTeamIdByName(accessToken, teamIdOrName);
}

/**
 * Resolve channelId - accepts either ID or name, and teamId or teamName
 */
async function resolveChannelId(accessToken, channelIdOrName, teamIdOrName) {
  if (!channelIdOrName) return null;
  if (!teamIdOrName) return null;
  
  // If channel looks like a GUID, assume it's already an ID
  if (channelIdOrName.match(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i)) {
    // Still need to resolve teamId if it's a name
    const teamId = await resolveTeamId(accessToken, teamIdOrName);
    return { channelId: channelIdOrName, teamId };
  }
  
  // Otherwise, treat channel as name and resolve
  return await findChannelIdByName(accessToken, channelIdOrName, teamIdOrName);
}

/**
 * Resolve userId - accepts email, UPN, displayName, or ID
 */
async function resolveUserId(accessToken, userIdOrEmailOrName) {
  if (!userIdOrEmailOrName) return null;
  
  // If it's a GUID, assume it's already an ID
  if (userIdOrEmailOrName.match(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i)) {
    return userIdOrEmailOrName;
  }
  
  // Otherwise, resolve from email or name
  return await findUserId(accessToken, userIdOrEmailOrName);
}

module.exports = {
  findTeamIdByName,
  findChannelIdByName,
  findUserId,
  findMeetingId,
  findChatId,
  resolveTeamId,
  resolveChannelId,
  resolveUserId
};


