# Microsoft Teams Graph API Tools

Node.js REST API server wrapping Microsoft Graph APIs for common Teams operations (messaging, team/channel provisioning, meeting transcripts/recordings, Copilot insights).

## Setup

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Create `.env` file:**
   ```env
   GRAPH_ACCESS_TOKEN=your_graph_access_token_here
   PORT=3000

   # For inbound alert webhook (Tool H1)
   TEAMS_INCOMING_WEBHOOK_URL=https://outlook.office.com/webhook/...
   INBOUND_WEBHOOK_SECRET=optional-shared-secret
   ```

3. **Start the server:**
   ```bash
   npm start
   ```

   Server runs on `http://localhost:3000`

## Authentication

You can provide the Graph access token in two ways:

1. **Authorization header** (recommended for Postman):
   ```
   Authorization: Bearer YOUR_ACCESS_TOKEN
   ```

2. **Request body** (for some endpoints):
   ```json
   {
     "accessToken": "YOUR_ACCESS_TOKEN",
     ...
   }
   ```

3. **Environment variable** (fallback):
   Set `GRAPH_ACCESS_TOKEN` in your `.env` file.

## Quick Start Guide

### Step 1: Get Your Access Token
You need a Microsoft Graph access token. You can get one from:
- Azure Portal (App Registration)
- Microsoft Graph Explorer
- Your authentication flow

### Step 2: Start the Server
```bash
npm start
```

### Step 3: Find IDs You Need
Before using the main endpoints, use the **Discovery endpoints** to find IDs:

1. **Get your user info:**
   ```
   GET http://localhost:3000/api/discovery/me
   ```

2. **List your teams:**
   ```
   GET http://localhost:3000/api/discovery/teams
   ```
   Copy a `teamId` from the response.

3. **List channels in a team:**
   ```
   GET http://localhost:3000/api/discovery/teams/{teamId}/channels
   ```
   Copy a `channelId` from the response.

4. **Find users:**
   ```
   GET http://localhost:3000/api/discovery/users?email=user@example.com
   ```
   Copy a `userId` (id field) from the response.

5. **List meetings:**
   ```
   GET http://localhost:3000/api/discovery/meetings
   ```
   Copy a `meetingId` from the response.

### Step 4: Use Postman Collection
1. Import `postman_collection.json` into Postman
2. Set the `accessToken` variable in the collection
3. Use Discovery requests to populate `teamId`, `channelId`, `userId`, etc.
4. Test the main endpoints!

## API Endpoints

### Discovery (Find IDs)

These endpoints help you discover the IDs needed for other operations.

#### GET `/api/discovery/me`
Get current user info.

#### GET `/api/discovery/teams`
List teams the user is a member of.
- Query param: `?all=true` to list all teams (requires admin)

#### GET `/api/discovery/teams/:teamId`
Get team details.

#### GET `/api/discovery/teams/:teamId/channels`
List channels in a team.

#### GET `/api/discovery/teams/:teamId/channels/:channelId`
Get channel details.

#### GET `/api/discovery/chats`
List chats the user is part of.

#### GET `/api/discovery/chats/:chatId`
Get chat details.

#### GET `/api/discovery/meetings`
List online meetings.
- Query params: `filter`, `top`

#### GET `/api/discovery/meetings/:meetingId`
Get meeting details.

#### GET `/api/discovery/users`
List users.
- Query params: `email`, `displayName`, `filter`, `search`, `top`

#### GET `/api/discovery/users/:userId`
Get user details.

#### GET `/api/discovery/teams/:teamId/members`
List team members.

#### GET `/api/discovery/teams/:teamId/channels/:channelId/members`
List channel members (for private/shared channels).

### Webhooks / Workflows

These tools do **not** require Graph auth; they rely on Teams webhook features.

#### Tool H1: Inbound alert webhook

- **Endpoint:** `POST /api/webhooks/inbound/alert`
- **Env:** `TEAMS_INCOMING_WEBHOOK_URL` (required), `INBOUND_WEBHOOK_SECRET` (optional)

**Request Headers (optional but recommended):**
```text
X-Shared-Secret: optional-shared-secret
Authorization: Bearer YOUR_ACCESS_TOKEN (not required for this endpoint)
```

**Request Body:**
```json
{
  "title": "Alert title (optional, default: Alert)",
  "summary": "Short description of the alert",
  "severity": "info | warning | critical",
  "source": "ci-cd-pipeline",
  "details": {
    "buildId": "123",
    "status": "failed"
  }
}
```

This endpoint forwards the payload to your Teams **incoming webhook URL** so you can plug it into CI/CD, monitoring, etc.

#### Tool H2: Outgoing webhook command bridge

- **Endpoint:** `POST /api/webhooks/outgoing`
- Configure this URL as an **Outgoing Webhook** in a Team.
- When users type `@YourWebhook some text`, Teams will POST here and expect a JSON response.

**Response Example:**
```json
{
  "text": "Hi John Doe, you said: \"some text\" in channel 19:abc123."
}
```

> Note: In production, you should validate the HMAC signature Teams sends to ensure the request is genuine.

### Messaging

#### POST `/api/messaging/channel/announcement`
Post an announcement to a channel.

**Request Body:**
```json
{
  "teamId": "team-id-here",
  "channelId": "channel-id-here",
  "content": "<p>Hello from API!</p>"
}
```

#### POST `/api/messaging/chat/send`
Send a DM in an existing chat.

**Request Body:**
```json
{
  "chatId": "chat-id-here",
  "content": "Hello!",
  "contentType": "html"
}
```

#### POST `/api/messaging/chat/create-and-send`
Create a new 1:1 or group chat, then send a message.

**Request Body:**
```json
{
  "chatType": "oneOnOne",
  "members": [
    { "id": "user-id-1" },
    { "id": "user-id-2" }
  ],
  "content": "Hello from new chat!",
  "contentType": "html"
}
```

#### POST `/api/messaging/subscriptions/create`
Create a chat change watcher subscription.

**Request Body:**
```json
{
  "resource": "/chats",
  "notificationUrl": "https://your-app.com/webhook",
  "expirationDateTime": "2024-12-31T23:59:59Z",
  "clientState": "optional-secret",
  "changeType": "created,updated"
}
```

### Teams & Channels

#### POST `/api/teams/create`
Create a new Team from a template.

**Request Body:**
```json
{
  "displayName": "Project Falcon",
  "description": "Team description",
  "templateId": "standard"
}
```

#### POST `/api/teams/:teamId/channels/create`
Create a channel in a team.

**URL Params:** `teamId`

**Request Body:**
```json
{
  "displayName": "Design",
  "description": "Design discussions",
  "membershipType": "standard"
}
```

#### POST `/api/teams/:teamId/members/add`
Add a member (or owner) to a team.

**URL Params:** `teamId`

**Request Body:**
```json
{
  "userId": "user-aad-object-id",
  "isOwner": false
}
```

#### POST `/api/teams/:teamId/channels/:channelId/members/add`
Add a member to a private/shared channel.

**URL Params:** `teamId`, `channelId`

**Request Body:**
```json
{
  "userId": "user-aad-object-id",
  "isOwner": false
}
```

#### POST `/api/teams/:teamId/tags/create`
Create a Teams tag.

**URL Params:** `teamId`

**Request Body:**
```json
{
  "displayName": "OnCall",
  "description": "On-call team members",
  "members": [
    { "userId": "user-id-1" },
    { "userId": "user-id-2" }
  ]
}
```

#### GET `/api/teams/:teamId/tags`
List Teams tags.

**URL Params:** `teamId`

### Meetings

#### GET `/api/meetings/:meetingId/transcripts`
List transcripts for a scheduled meeting.

**URL Params:** `meetingId`

#### GET `/api/meetings/:meetingId/transcripts/:transcriptId`
Get transcript metadata.

**URL Params:** `meetingId`, `transcriptId`

#### GET `/api/meetings/:meetingId/transcripts/:transcriptId/content`
Get transcript content (VTT file).

**URL Params:** `meetingId`, `transcriptId`

#### GET `/api/meetings/:meetingId/recordings`
List recordings for a scheduled meeting.

**URL Params:** `meetingId`

#### GET `/api/meetings/:meetingId/ai-insights`
List AI insights for a meeting (filtered by type).

**URL Params:** `meetingId`

**Query Params:**
- `userId` (required)
- `aiInsightType` (required): `meetingSummary`, `callToAction`, or `mention`

**Example:** `/api/meetings/{meetingId}/ai-insights?userId={userId}&aiInsightType=meetingSummary`

#### GET `/api/meetings/:meetingId/ai-insights/:aiInsightId`
Get a specific AI insight.

**URL Params:** `meetingId`, `aiInsightId`

**Query Params:**
- `userId` (required)

**Example:** `/api/meetings/{meetingId}/ai-insights/{aiInsightId}?userId={userId}`

## Postman Collection

### Import the Collection

1. **Open Postman** and click **Import**
2. **Select** `postman_collection.json` from this project
3. **Set Collection Variables:**
   - Click on the collection name
   - Go to **Variables** tab
   - Set `accessToken` to your Graph access token
   - Set `baseUrl` if different from `http://localhost:3000`

### Using the Collection

1. **Start with Discovery:**
   - Run `Get Current User` to verify your token works
   - Run `List My Teams` to get team IDs
   - Run `List Channels in Team` (set `teamId` variable first)
   - Run `List Users` or `Search Users by Email` to find user IDs
   - Run `List Online Meetings` to find meeting IDs

2. **Use Variables:**
   - After running discovery requests, copy IDs from responses
   - Set collection variables: `teamId`, `channelId`, `userId`, `meetingId`, etc.
   - These will auto-populate in other requests

3. **Test Main Endpoints:**
   - All requests use the `Authorization` header automatically
   - Just update the variables and send requests!

### Manual Setup (Alternative)

1. **Base URL:** `http://localhost:3000`

2. **Authorization:** Add to Headers:
   - Key: `Authorization`
   - Value: `Bearer YOUR_ACCESS_TOKEN`

3. **Content-Type:** `application/json` (automatically set)

### Example Postman Requests

#### Post Announcement
```
POST http://localhost:3000/api/messaging/channel/announcement
Headers:
  Authorization: Bearer YOUR_TOKEN
Body (JSON):
{
  "teamId": "your-team-id",
  "channelId": "your-channel-id",
  "content": "<p>Test announcement</p>"
}
```

#### Create Team
```
POST http://localhost:3000/api/teams/create
Headers:
  Authorization: Bearer YOUR_TOKEN
Body (JSON):
{
  "displayName": "My New Team",
  "description": "Team description",
  "templateId": "standard"
}
```

#### List Meeting Transcripts
```
GET http://localhost:3000/api/meetings/{meetingId}/transcripts
Headers:
  Authorization: Bearer YOUR_TOKEN
```

**Note:** Get `meetingId` from `GET /api/discovery/meetings` first!

## Notes

- **Copilot AI Insights:** Requires a licensed Copilot user and meetings where transcription/recording are enabled. Insights may take up to 4 hours after meeting end.
- **Team Creation:** For production, follow Microsoft's recommended approach (create backing group first, then convert to team).
- **Rate Limiting:** Microsoft recommends staggering team member addition calls (2-second buffer mentioned in docs).

