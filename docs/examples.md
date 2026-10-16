# Usage examples

Each example pairs a natural-language prompt with the tool call an assistant makes. Arguments are the tool's JSON input; every tool is documented in the [Tool reference](tools.md#tool-reference).

## Session — `session_status`, `auth_begin_login`

> "Am I signed in?"

```json
{}
```

`session_status` with no arguments reports the cached account plus live `GET /me` reachability. Call it once first. If no session is signed in:

```json
{ "mode": "full" }
```

`auth_begin_login` opens the system browser and returns immediately; the user finishes there and the assistant confirms with `auth_login_status`.

## Mail — `mail_search_messages`, `mail_fetch_thread`

> "Summarize the reply chain about the Q3 invoice."

1. Search first. A date-bounded search uses `$filter` and matches the subject:

```json
{ "query": "Q3 invoice", "after": "2026-09-01" }
```

2. Read the whole thread in one call — never loop `mail_get_message` over a reply chain:

```json
{ "fromMessageId": "<id from step 1>" }
```

`mail_fetch_thread` folds text the messages repeat at each other into `[quoted from #N]` markers.

## Calendar — `calendar_find_meeting_times`, `calendar_create_event`

> "Find a 30-minute slot with Ana and Ben next week, then book it."

```json
{
  "attendees": ["ana@contoso.com", "ben@contoso.com"],
  "durationMinutes": 30,
  "start": "2026-10-19T09:00:00",
  "end": "2026-10-23T18:00:00",
  "timeZone": "UTC"
}
```

Then create on a suggested slot:

```json
{
  "subject": "Q3 review",
  "start": "2026-10-21T14:00:00",
  "end": "2026-10-21T14:30:00",
  "attendees": [{ "address": "ana@contoso.com" }, { "address": "ben@contoso.com" }],
  "isOnlineMeeting": true
}
```

Note the shapes differ: `calendar_find_meeting_times` takes plain address strings, while `calendar_create_event` takes `{ "address" }` objects.

## Files — `files_search`, `files_parse_document`

> "What does the shared Q3 report say about risks?"

```json
{ "query": "Q3 report risks" }
```

Then extract the document as unit rows (sections/pages) from a sharing link:

```json
{ "shareUrl": "https://contoso-my.sharepoint.com/:w:/g/personal/..." }
```

`files_parse_document` also accepts a `driveId` + `itemId` pair or a `messageId` + `attachmentId` pair, and pages long documents with `unitOffset` / `unitLimit` or the manifest `cursor`.

## To Do — `todo_list_lists`, `todo_create_task`

> "Remind me to file expenses on Friday."

```json
{ "displayName": "Personal" }
```

is `todo_create_list` when no list fits; otherwise create directly:

```json
{
  "listId": "<id from todo_list_lists>",
  "title": "File expenses",
  "dueDateTime": "2026-10-23",
  "importance": "high"
}
```

## Planner — `planner_list_my_tasks`, `planner_create_task`

> "What is assigned to me, and add a follow-up to the Launch plan."

```json
{ "limit": 20 }
```

```json
{
  "planId": "<plan id>",
  "title": "Follow up on launch checklist",
  "dueDateTime": "2026-10-30T17:00:00Z"
}
```

Updates and deletes supply their own `If-Match` etag — no extra read needed.

## Contacts — `contacts_search`

> "Find Priya's email address."

```json
{ "query": "Priya", "top": 5 }
```

## Chat — `chat_recent_activity`, `chat_fetch_history`

> "What did I miss in Teams yesterday?"

```json
{ "chats": 10, "messagesPerChat": 3 }
```

`chat_recent_activity` returns the most-active chats first with a preview each. For depth on one chat:

```json
{ "chatId": "<chat id>", "since": "2026-10-15", "maxMessages": 200 }
```

## SharePoint — `sharepoint_search_sites`, `sharepoint_list_list_items`

> "Show open items on the Engineering tracker."

```json
{ "query": "Engineering" }
```

```json
{
  "siteId": "<site id>",
  "listId": "<list id>",
  "filter": "Status ne 'Closed'",
  "top": 25
}
```

## Search — `search_query`

> "Find everything about the Fabrikam migration."

```json
{ "query": "Fabrikam migration", "entityTypes": ["message", "driveItem", "site"] }
```

One call queries Microsoft's index across mail, files and sites. Add `"chatMessage"` only with the `teams` group enabled and admin consent — see [Search permissions](auth.md#search-permissions).

## Teams — `teams_fetch_channel_history`

> "Catch me up on #general this week, including replies."

```json
{
  "teamId": "<team id>",
  "channelId": "<channel id>",
  "since": "2026-10-12",
  "includeReplies": true,
  "maxMessages": 100
}
```

Channel queries cannot filter server-side, so the tool pages and filters for you.

## Directory — `directory_search_users`

> "Who is Ana's manager?"

```json
{ "query": "Ana", "top": 5 }
```

```json
{ "id": "<user id from step 1>" }
```

The second call is `directory_get_user_manager`. Needs `--org-mode` and admin consent.

## Intune — `intune_list_noncompliant`

> "Which devices are noncompliant?"

```json
{ "top": 25 }
```

Needs `--org-mode` and admin consent.

## Generic — `graph_request`, `graph_schema`

> "List my joined teams with only id and name."

```json
{
  "path": "/me/joinedTeams",
  "query": { "$select": "id,displayName", "$top": "25" }
}
```

For an unfamiliar endpoint, inspect its shape first:

```json
{ "entityPath": "/me/joinedTeams" }
```

`graph_schema` fetches one item and reports its property names and types. Writes through `graph_request` need `--allow-generic-write`.

## Discovery mode — `discover_tools`, `call_tool`

> "Which tool reads attachments?"

```json
{ "query": "attachment" }
```

`discover_tools` searches names and descriptions without loading schemas. Then invoke by exact name:

```json
{ "name": "mail_list_attachments", "arguments": { "id": "<message id>" } }
```
