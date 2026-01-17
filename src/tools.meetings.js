const { GraphClient } = require("./graphClient");

/**
 * Meeting tools: transcripts, recordings, and Copilot AI insights.
 */

/**
 * Tool: "Pull the transcript from a scheduled meeting"
 *
 * List callTranscript objects for an onlineMeeting.
 * (Exact resource path may vary by API version/shape; this matches current docs style.)
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.meetingId - onlineMeeting ID associated with a calendar event.
 */
async function listMeetingTranscripts(accessToken, { meetingId }) {
  const client = new GraphClient(accessToken);
  return client.get(`/me/onlineMeetings/${meetingId}/transcripts`);
}

/**
 * Get transcript metadata.
 * API: GET /me/onlineMeetings/{meetingId}/transcripts/{transcriptId}
 */
async function getMeetingTranscript(accessToken, { meetingId, transcriptId }) {
  const client = new GraphClient(accessToken);
  return client.get(`/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}`);
}

/**
 * Get transcript content (e.g. VTT).
 * API: GET /me/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content
 */
async function getMeetingTranscriptContent(accessToken, { meetingId, transcriptId }) {
  const client = new GraphClient(accessToken);
  // Expect binary/text content; override responseType when needed.
  return client.get(
    `/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`,
    { responseType: "arraybuffer" }
  );
}

/**
 * Tool: "Pull the recording list for a scheduled meeting"
 * (Shape is similar to transcripts, depending on current Graph documentation.)
 */
async function listMeetingRecordings(accessToken, { meetingId }) {
  const client = new GraphClient(accessToken);
  return client.get(`/me/onlineMeetings/${meetingId}/recordings`);
}

/**
 * Tool: "Generate meeting minutes from Copilot AI insights"
 * Meeting AI insights APIs:
 *  - GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights/{aiInsightId}
 *  - GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights?$filter=aiInsightType eq 'meetingSummary'
 *  - GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights?$filter=aiInsightType eq 'callToAction'
 *  - GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights?$filter=aiInsightType eq 'mention'
 *
 * NOTE: Requires a licensed Copilot user and meetings where transcription/recording are enabled.
 */

/**
 * List AI insights of a specific type for a meeting.
 *
 * @param {string} accessToken
 * @param {object} params
 * @param {string} params.userId
 * @param {string} params.meetingId
 * @param {"meetingSummary"|"callToAction"|"mention"} params.aiInsightType
 */
async function listMeetingAiInsights(accessToken, { userId, meetingId }) {
  const client = new GraphClient(accessToken);

  return client.get(
    `/copilot/users/${userId}/onlineMeetings/${meetingId}/aiInsights`
  );
}

/**
 * Get a specific AI insight object.
 */
async function getMeetingAiInsight(accessToken, { userId, meetingId, aiInsightId }) {
  const client = new GraphClient(accessToken);
  return client.get(
    `/copilot/users/${userId}/onlineMeetings/${meetingId}/aiInsights/${aiInsightId}`
  );
}

module.exports = {
  listMeetingTranscripts,
  getMeetingTranscript,
  getMeetingTranscriptContent,
  listMeetingRecordings,
  listMeetingAiInsights,
  getMeetingAiInsight
};


