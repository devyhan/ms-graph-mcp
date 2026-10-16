# Tools

111 tools across 14 groups. `--read-only` removes every write tool, leaving 80. The reference for each group is below; the descriptions here are the first line of what the model itself is shown.

## Tool groups

Tools are grouped by product area. Enable only the groups you need — fewer tools, less surface. Sign-in is incremental: a bare login asks only for the minimal set, and consent grows as tools are first used.

| Group | Product | Delegated scopes | Admin consent | Tools |
| --- | --- | --- | --- | --- |
| `me` | Entra ID profile, Outlook mailbox settings | `User.Read`, `MailboxSettings.Read` | No | 2 |
| `mail` | Outlook mail | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` | No | 14 |
| `calendar` | Outlook calendar | `Calendars.Read`, `Calendars.Read.Shared`, `Calendars.ReadWrite` | No | 9 |
| `files` | OneDrive | `Files.Read`, `Files.Read.All`, `Files.ReadWrite`, `Mail.Read` ² | No | 13 |
| `todo` | Microsoft To Do | `Tasks.Read`, `Tasks.ReadWrite` | No | 8 |
| `planner` | Microsoft Planner | `Tasks.Read`, `Tasks.ReadWrite` | No | 9 |
| `contacts` | Outlook contacts | `Contacts.Read`, `Contacts.ReadWrite` | No | 6 |
| `chat` | Teams one-to-one and group chats | `Chat.Read`, `Chat.ReadWrite` | No | 7 |
| `sharepoint` | SharePoint sites, libraries, lists | `Sites.Read.All`, `Sites.ReadWrite.All` | No | 11 |
| `search` | Microsoft Search across mail, files, sites | `Mail.Read`, `Files.Read`, `Sites.Read.All`, `Calendars.Read`, `Chat.Read` | No¹ | 1 |
| `teams` | Teams teams, channels, channel messages | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `TeamMember.Read.All` | **Yes** | 10 |
| `directory` | Entra ID user and group lookups | `User.Read.All`, `Group.Read.All` | **Yes** | 9 |
| `intune` | Intune managed device inventory | `DeviceManagementManagedDevices.Read.All` | **Yes** | 6 |
| `generic` | Signing in, any Graph endpoint, introspection | reuses the union of the above | No | 6 |

`offline_access` and `User.Read` are always requested: the first keeps refresh tokens working, the second identifies the account.

The three groups marked **Yes** request scopes that no ordinary user can consent to, so they stay disabled unless you pass `--org-mode`. Enabling them in a tenant where an administrator has not granted consent means sign-in fails outright, not that those tools quietly return errors.

The **Admin consent** column is Microsoft's own classification of each permission. A tenant's consent policy is a second, separate gate, and it withholds more than this column shows — `Mail.Read` and `Calendars.Read` among them. See [Consent](auth.md#consent).

¹ `search` itself needs no administrator, but one thing it can search does: the `chatMessage` index additionally requires `ChannelMessage.Read.All`, which belongs to `teams`. See [Search permissions](auth.md#search-permissions).

² `Mail.Read` rides on the one files tool that reads mail attachments, `files_parse_document`. Consent is incremental: a drive-only caller never pays for it.

The `generic` group is always on and cannot be selected or removed. It holds `auth_begin_login` and `auth_login_status` (the sign-in pair), `session_status` (the session probe), `graph_request` (the escape hatch for endpoints no other tool covers), `graph_schema` (fetches one item from a path and reports its property names and types), and `graph_list_permissions` (reports the enabled groups, their scopes, and the signed-in account).

## Tool reference

<details>
<summary><b><code>me</code></b> — My profile: 2 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `me_get_profile` |  | Returns the signed-in user's own Entra ID profile from /me: by default id, displayName, mail, userPrincipalName, jobTitle, officeLocation and preferredLanguage. | `select` |
| `me_get_mailbox_settings` |  | Returns the signed-in user's Outlook mailbox configuration: time zone, date and time format, locale, working hours, and the automatic-reply (out-of-office) setting. | `replyBodyChars` |

</details>

<details>
<summary><b><code>mail</code></b> — Mail: 14 tools, 7 of them writes</summary>

Bodies are the expensive part of a mailbox, so list calls return `bodyPreview` only and `mail_get_message` fetches one message in full, HTML stripped to text by default.

`mail_list_messages` and `mail_search_messages` no longer return `webLink` unless you ask for it with `includeWebLink`. An Outlook webLink is the message id a second time — URL-encoded and wrapped in an owa query — and on a real 217-row search it was 22% of the whole response. The `id` is what every other mail tool takes, so nothing you could do before is harder; naming `webLink` in `select` still works too.

A reply chain is the expensive case, because Outlook quotes itself: message N carries most of messages 1..N-1 inside its own body, so reading a thread one call at a time sends the same paragraphs through the model's context once per reply. `mail_fetch_thread` reads the whole chain in one call and replaces each repeat with a marker naming the message that carried it first. On a measured 24-message Outlook thread that was 381,198 characters through `mail_get_message` and 85,171 through `mail_fetch_thread`, with every unique sentence preserved.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `mail_list_messages` |  | Lists Outlook messages newest-first, 10 per call by default (50 max), as a compact projection: id, subject, from, toRecipients, receivedDateTime, isRead, hasAttachments and bodyPreview. The Outlook `webLink` is off by default — see below. | `folderId`, `top`, `skip`, `filter`, `orderby`, `select`, `unreadOnly`, `includeWebLink` |
| `mail_search_messages` |  | Finds messages by text, returning the same compact projection as mail_list_messages, 10 per call by default (50 max). | **`query`**, `top`, `from`, `to`, `after`, `before`, `hasAttachments`, `includeWebLink` |
| `mail_get_message` |  | Returns one message with its body plus sender, recipients, timestamps, importance, conversationId and webLink. Do not loop it over a reply chain — use `mail_fetch_thread`. | **`id`**, `format`, `maxBodyChars`, `foldLinks` |
| `mail_fetch_thread` |  | Reads a whole reply chain in one call, oldest-first, folding text the messages repeat at each other into markers like `[quoted from #3]`. Give it any message id of the thread, or an explicit list of ids. | `fromMessageId`, `ids`, `maxMessages`, `maxBodyChars`, `totalBodyChars`, `fold`, `foldLinks` |
| `mail_list_folders` |  | Lists the top-level Outlook mail folders with their ids, unread counts and total counts, 50 per call by default. | `top`, `includeChildren` |
| `mail_list_attachments` |  | Lists attachment metadata for one message — id, name, contentType, size in bytes, whether it is inline, and its kind (file, item, or reference) — 20 per call by default. | **`id`**, `top` |
| `mail_get_attachment_content` |  | Returns one attachment's content: text for text-like files under 1 MB, base64 otherwise, or the sourceUrl for a reference attachment. | **`id`**, **`attachmentId`**, `maxChars` |
| `mail_send` | **write** | Sends a message immediately from the signed-in mailbox — there is no undo and no confirmation step, so confirm the recipients and text with the user first. | **`to`**, **`subject`**, **`body`**, `cc`, `bcc`, `contentType`, `saveToSentItems` |
| `mail_create_draft` | write | Creates an unsent draft in the Drafts folder and returns its id and webLink. | `to`, **`subject`**, **`body`**, `cc`, `bcc`, `contentType` |
| `mail_reply` | **write** | Sends a reply to an existing message immediately — it does not create a draft, and there is no undo, so confirm the text with the user first. | **`id`**, **`comment`**, `replyAll` |
| `mail_forward` | **write** | Forwards an existing message immediately, attachments included — it does not create a draft and there is no undo, so confirm the recipients with the user first. | **`id`**, **`to`**, `comment` |
| `mail_move_message` | write | Moves a message to another mail folder. | **`id`**, **`destinationId`** |
| `mail_mark_read` | write | Sets the read flag on one message. | **`id`**, `isRead` |
| `mail_delete_message` | **write** | Deletes a message. | **`id`** |

</details>

<details>
<summary><b><code>calendar</code></b> — Calendar: 9 tools, 4 of them writes</summary>

`calendar_list_events` switches to `/me/calendarView` when you give it a date range, which is the only form that expands recurring series. `calendar_get_schedule` and `calendar_find_meeting_times` are POSTs that read rather than write.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `calendar_list_events` |  | Lists events with subject, start, end, location, organizer and attendee count, plus a nextLink when more pages exist. | `start`, `end`, `calendarId`, `top`, `select`, `orderby`, `timeZone` |
| `calendar_get_event` |  | Returns one event in full: subject, start, end, location, organizer, the attendee list with each response, the description body (HTML converted to text and capped), recurrence summary, and the Outlook web link. | **`id`**, `timeZone` |
| `calendar_list_calendars` |  | Lists the signed-in user's calendars with id, name, colour, owner and whether they can be edited or shared. | `top` |
| `calendar_get_schedule` |  | Returns free/busy availability for up to 20 people or rooms over a time window, as an availabilityView string where each character covers one interval (0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere), plus working hours and any visible busy blocks. | **`schedules`**, **`start`**, **`end`**, `availabilityViewInterval`, `timeZone` |
| `calendar_find_meeting_times` |  | Suggests meeting slots that work for the signed-in user and the given attendees, ranked by confidence, with the attendees who are unavailable in each slot. | **`attendees`**, **`durationMinutes`**, `start`, `end`, `maxCandidates`, `timeZone` |
| `calendar_create_event` | write | Creates an event on the primary calendar and returns the created event. | **`subject`**, **`start`**, **`end`**, `timeZone`, `attendees`, `body`, `bodyType`, `location`, `isAllDay`, `isOnlineMeeting` |
| `calendar_update_event` | write | Patches the given fields on an event and returns the updated event; omitted fields are left alone. | **`id`**, `subject`, `start`, `end`, `timeZone`, `attendees`, `body`, `bodyType`, `location`, `isAllDay`, `isOnlineMeeting`, `showAs`, `categories`, `reminderMinutesBeforeStart` |
| `calendar_delete_event` | **write** | Deletes an event, moving it to Deleted Items. | **`id`** |
| `calendar_respond_event` | **write** | Accepts, declines, or tentatively accepts a meeting invitation, optionally with a comment to the organizer. | **`id`**, **`response`**, `comment`, `sendResponse` |

</details>

<details>
<summary><b><code>files</code></b> — Files: 13 tools, 4 of them writes</summary>

Path-addressed calls encode each segment separately, so spaces and `&` in filenames work. `files_read_text` refuses binary types by extension before downloading, and no tool returns `@microsoft.graph.downloadUrl`, which is a pre-authenticated link to the file contents.

Documents referenced from elsewhere — a Teams attachment `contentUrl`, a channel Files tab, a mail attachment — reach the model through `files_parse_document`, which extracts `.docx`/`.xlsx`/`.pptx`/`.pdf` as unit rows (sections/sheets/slides/pages) with a manifest. It is backed by five pure-JS libraries, the only new runtime dependencies: mammoth (structure-aware docx), exceljs (streaming xlsx; the frozen `xlsx` package has known CVEs), jszip plus fast-xml-parser (ordered pptx slides), and pdfjs-dist (positioned PDF text).

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `files_list_children` |  | Lists the direct children of a folder in the signed-in user's OneDrive, returning id, name, kind (file or folder), size, mime type, child count, last-modified time and webUrl per entry, plus count and nextLink when more pages exist. | `itemId`, `path`, `top`, `orderby` |
| `files_get_item` |  | Reads the metadata of one OneDrive file or folder: name, kind, size, mime type, created and last-modified times and authors, parent folder path and ids, and webUrl. | `itemId`, `path` |
| `files_search` |  | Searches the whole of the signed-in user's OneDrive for free text, matching file names and indexed file contents, and returns the same compact fields as files_list_children plus count and nextLink. | **`query`**, `top` |
| `files_list_recent` |  | Lists files the signed-in user recently opened or edited, most recent first, with count and nextLink. | `top` |
| `files_list_shared` |  | Lists files and folders other people have shared with the signed-in user, with count and nextLink. | `top` |
| `files_read_text` |  | Returns the contents of a text-like OneDrive file (plain text, Markdown, CSV, JSON, XML, source code, config) as a string, truncated to maxChars, which defaults to 20000. | **`itemId`**, `maxChars` |
| `files_resolve_share` |  | Resolves a OneDrive/SharePoint sharing link to the drive address behind it: driveId plus itemId, name, kind, size, mime type and webUrl. | **`shareUrl`** |
| `files_get_drive_item` |  | Reads one file or folder in any reachable drive (shared files, channel folders), returning the same detail as files_get_item. | **`driveId`**, **`itemId`** |
| `files_parse_document` |  | Extracts a .docx, .xlsx, .pptx or .pdf as unit rows (sections/sheets/slides/pages) plus a manifest with unitCount, per-unit truncation flags and a cursor. | `shareUrl`, `driveId`+`itemId`, `messageId`+`attachmentId`, `unitOffset`, `unitLimit`, `maxCharsPerUnit`, `cursor` |
| `files_create_folder` | write | Creates a folder in the signed-in user's OneDrive and returns the created item. | **`name`**, `parentItemId`, `parentPath` |
| `files_upload_text` | write | Writes a text file to the given path in the signed-in user's OneDrive and returns the stored item. | **`path`**, **`content`** |
| `files_delete` | **write** | Deletes a file or folder from the signed-in user's OneDrive. | **`itemId`** |
| `files_create_link` | write | Creates a sharing link for a OneDrive file or folder and returns its URL, permission type and scope. | **`itemId`**, `type`, `scope` |

</details>

<details>
<summary><b><code>todo</code></b> — To Do: 8 tools, 5 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `todo_list_lists` |  | Lists the signed-in user's Microsoft To Do lists with their id, display name and well-known name (for example defaultList for the built-in Tasks list). | none |
| `todo_list_tasks` |  | Lists tasks in one To Do list, returning id, title, status, importance, due and reminder times and a short body preview per task, plus count and nextLink when more pages exist. | **`listId`**, `top`, `filter`, `includeCompleted` |
| `todo_get_task` |  | Reads one To Do task in full: title, status, importance, reminder state, start/due/completed times, categories and the notes body (truncated). | **`listId`**, **`taskId`** |
| `todo_create_task` | write | Creates a task in a To Do list and returns the created task. | **`listId`**, **`title`**, `body`, `dueDateTime`, `reminderDateTime`, `importance`, `timeZone` |
| `todo_update_task` | write | Patches an existing To Do task and returns the updated task. | **`listId`**, **`taskId`**, `title`, `body`, `status`, `importance`, `dueDateTime`, `reminderDateTime`, `clearDueDateTime`, `clearReminderDateTime`, `timeZone` |
| `todo_complete_task` | write | Marks a To Do task completed by patching its status to 'completed', and returns the updated task including the completedDateTime Graph stamps on it. | **`listId`**, **`taskId`** |
| `todo_delete_task` | **write** | Deletes a To Do task permanently and returns a confirmation. | **`listId`**, **`taskId`** |
| `todo_create_list` | write | Creates a new To Do list and returns its id and display name. | **`displayName`** |

</details>

<details>
<summary><b><code>planner</code></b> — Planner: 9 tools, 3 of them writes</summary>

Planner requires an `If-Match` etag on every update and delete. The tools fetch the item first and supply the etag themselves.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `planner_list_my_plans` |  | Lists the Planner plans the signed-in user can see, with id, title, owning Microsoft 365 group id and container. | `limit` |
| `planner_list_my_tasks` |  | Lists Planner tasks assigned to the signed-in user across every plan, with title, planId, bucketId, due date, percentComplete and assignee ids. | `limit` |
| `planner_get_plan` |  | Returns one Planner plan: id, title, owning Microsoft 365 group id, container and creation time. | **`planId`** |
| `planner_list_buckets` |  | Lists the buckets (columns) of a plan with id, name and orderHint. | **`planId`**, `limit` |
| `planner_list_plan_tasks` |  | Lists every task in a plan regardless of assignee, with title, bucketId, due date, percentComplete and assignee ids. | **`planId`**, `limit` |
| `planner_get_task` |  | Returns one Planner task. | **`taskId`**, `includeDetails` |
| `planner_create_task` | write | Creates a task in a plan and returns it. | **`planId`**, **`title`**, `bucketId`, `assigneeIds`, `dueDateTime`, `percentComplete`, `priority` |
| `planner_update_task` | write | Updates a Planner task and returns it. | **`taskId`**, `title`, `bucketId`, `assigneeIds`, `dueDateTime`, `clearDueDateTime`, `percentComplete`, `priority` |
| `planner_delete_task` | **write** | Permanently deletes a Planner task; there is no recycle bin. | **`taskId`** |

</details>

<details>
<summary><b><code>contacts</code></b> — Contacts: 6 tools, 3 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `contacts_list` |  | Lists personal contacts from the signed-in user's default Outlook Contacts folder, 25 per page by default (max 100). | `top`, `skip`, `filter`, `orderby`, `select` |
| `contacts_get` |  | Returns one personal contact by id, with names, company, job title, department, every email address, phone numbers, postal addresses, birthday, categories and timestamps. | **`id`** |
| `contacts_search` |  | Full-text search over the default Contacts folder using Outlook `$search`, returning the same compact rows as contacts_list, 25 per page by default. | **`query`**, `top`, `skip` |
| `contacts_create` | write | Creates a personal contact in the signed-in user's default Contacts folder and returns the stored contact, including the new id. | **`givenName`**, `surname`, `emailAddresses`, `mobilePhone`, `companyName`, `jobTitle` |
| `contacts_update` | write | Updates the given fields of one personal contact and returns the stored result. | **`id`**, `givenName`, `surname`, `displayName`, `emailAddresses`, `mobilePhone`, `companyName`, `jobTitle` |
| `contacts_delete` | **write** | Deletes one personal contact. | **`id`** |

</details>

<details>
<summary><b><code>chat</code></b> — Chat: 7 tools, 1 of them writes</summary>

Graph offers no `$select` and no date filter on chat messages, so projection happens here and date windows are applied while paging. `chat_fetch_history` is the one to reach for beyond a single page.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `chat_list_chats` |  | Lists the Teams chats the signed-in user belongs to (one-to-one, group, and meeting chats — never channel conversations, which the teams group covers). | `top`, `filter`, `expandMembers` |
| `chat_get_chat` |  | Returns one Teams chat by id: topic, chatType, created and last-updated timestamps, the Teams deep link, whether the signed-in user has hidden it, how far they have read, and the meeting join URL for a meeting chat. | **`chatId`** |
| `chat_list_messages` |  | Reads recent messages from one Teams chat, newest first. | **`chatId`**, `top`, `orderby` |
| `chat_list_members` |  | Lists everyone in one Teams chat: the per-chat membership id, display name, email, directory user id, and roles (an owner of a group chat has ["owner"]; ordinary members have none). | **`chatId`** |
| `chat_send_message` | **write** | Posts a message to a Teams chat as the signed-in user. | **`chatId`**, **`content`**, `contentType` |
| `chat_fetch_history` |  | Reads a long run of one Teams chat, newest first, paging automatically so you do not have to follow nextLink yourself. | **`chatId`**, `since`, `until`, `maxMessages`, `bodyChars`, `includeSystem`, `cursor` |
| `chat_recent_activity` |  | Lists the most-active Teams chats first with a last-message preview each, server-ordered in one request; set messagesPerChat above 1 to include recent messages per chat via one batched fetch. | `chats`, `messagesPerChat`, `bodyChars`, `includeSystem` |

</details>

<details>
<summary><b><code>sharepoint</code></b> — SharePoint: 11 tools, 3 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `sharepoint_search_sites` |  | Searches SharePoint sites and returns id, display name, description and web URL for each match, plus a count and a nextLink when more pages exist. | **`query`**, `top` |
| `sharepoint_get_site` |  | Returns one site: its composite id, display name, description, web URL, hostname and timestamps. | **`siteId`** |
| `sharepoint_list_drives` |  | Lists the document libraries (drives) on a site with id, name, description, drive type and web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, `top` |
| `sharepoint_list_drive_items` |  | Lists the children of a document-library folder: name, whether it is a file or a folder, size, child count, MIME type, last modified time and author, and the web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, **`driveId`**, `itemId`, `folderPath`, `top` |
| `sharepoint_list_lists` |  | Lists the SharePoint lists on a site with id, display name, description, template and web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, `includeHidden`, `top` |
| `sharepoint_get_list` |  | Returns one list: id, display name, description, template, web URL and timestamps. | **`siteId`**, **`listId`**, `includeColumns` |
| `sharepoint_list_list_items` |  | Lists items in a SharePoint list with their column values, plus a count and a nextLink when more pages exist. | **`siteId`**, **`listId`**, `top`, `filter`, `fields` |
| `sharepoint_get_list_item` |  | Returns one list item with its column values, creator, last editor, timestamps and web URL. | **`siteId`**, **`listId`**, **`itemId`**, `fields` |
| `sharepoint_create_list_item` | write | Creates an item in a SharePoint list and returns its new id, web URL and stored column values. | **`siteId`**, **`listId`**, **`fields`** |
| `sharepoint_update_list_item` | write | Updates column values on an existing list item and returns the stored values. | **`siteId`**, **`listId`**, **`itemId`**, **`fields`** |
| `sharepoint_delete_list_item` | **write** | Deletes one item from a SharePoint list. | **`siteId`**, **`listId`**, **`itemId`** |

</details>

<details>
<summary><b><code>search</code></b> — Search: 1 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `search_query` |  | Runs a relevance-ranked Microsoft Search query across Microsoft 365 and returns a flattened list of hits, each with its rank, a text summary and the resource's id, name or subject, web URL and last-modified time. Each `entityTypes` value asks only for the scope its own index needs — see [Search permissions](auth.md#search-permissions). | **`query`**, `entityTypes`, `from`, `size`, `fields` |

</details>

<details>
<summary><b><code>teams</code></b> — Teams: 10 tools, 1 of them writes</summary>

Channel messages have an even narrower query surface than chats: no `$filter`, no `$orderby`, and an order based on the reply chain rather than creation time. `teams_fetch_channel_history` filters client-side for that reason and says so.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `teams_list_joined` |  | Lists the Microsoft Teams teams the signed-in user is a member of, returning id, displayName, description, visibility and archived state. | none |
| `teams_list_channels` |  | Lists the channels of one team, returning id, displayName, description, membershipType and webUrl for each. | **`teamId`**, `membershipType` |
| `teams_get_channel` |  | Returns one channel: displayName, description, membershipType, creation time, the channel email address (empty unless the channel has one provisioned) and the deep link webUrl. | **`teamId`**, **`channelId`** |
| `teams_get_channel_files_folder` |  | Returns the SharePoint folder behind a channel's Files tab as a drive address (driveId plus itemId) for files_get_drive_item and files_parse_document. | **`teamId`**, **`channelId`** |
| `teams_list_channel_messages` |  | Lists the top-level messages of a channel, newest first, with sender, timestamp, plain-text body (HTML stripped and truncated to `bodyChars`) and attachment, mention and reaction counts. | **`teamId`**, **`channelId`**, `top`, `bodyChars` |
| `teams_get_channel_message` |  | Returns one channel message in full: sender, timestamps, plain-text body truncated to `bodyChars`, attachment names, mention text and reaction types. | **`teamId`**, **`channelId`**, **`messageId`**, `bodyChars` |
| `teams_list_message_replies` |  | Lists the replies to one top-level channel message, newest first, with the same compact projection as teams_list_channel_messages. | **`teamId`**, **`channelId`**, **`messageId`**, `top`, `bodyChars` |
| `teams_fetch_channel_history` |  | Walks a channel's message history and returns a bounded page of root messages with sender, creation and modification times, plain-text body (HTML stripped, truncated to `bodyChars`) and attachment count. | **`teamId`**, **`channelId`**, `since`, `until`, `maxMessages`, `bodyChars`, `includeReplies`, `cursor` |
| `teams_list_members` |  | Lists the members of a team with displayName, email, the directory userId and their roles ("owner" for owners, empty for ordinary members, "guest" for guests). | **`teamId`**, `top` |
| `teams_send_channel_message` | **write** | Posts a new top-level message to a channel as the signed-in user and returns the created message id, timestamp and deep link. | **`teamId`**, **`channelId`**, **`content`**, `contentType`, `subject`, `importance` |

</details>

<details>
<summary><b><code>directory</code></b> — Directory: 9 tools, all read-only</summary>

Advanced queries against the directory need the `ConsistencyLevel: eventual` header; the client sets it automatically when a request uses `$count` or a directory `$search`.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `directory_list_users` |  | Lists users in the Microsoft Entra ID tenant, 25 per page by default (max 999). | `top`, `filter`, `select`, `orderby`, `skipToken` |
| `directory_search_users` |  | Finds users by name or email address, returning the same compact rows as directory_list_users, 25 per page by default. | **`query`**, `top`, `select`, `skipToken` |
| `directory_get_user` |  | Returns one user by object id or userPrincipalName, with names, contact details, job title, department, office, employee id, account state, usage location and on-premises sync fields. | **`id`** |
| `directory_get_user_manager` |  | Returns the user's manager as a compact row (id, displayName, userPrincipalName, mail, jobTitle, department). | **`id`** |
| `directory_list_direct_reports` |  | Lists the users who report directly to the given user, 25 per page by default (max 999), as compact rows. | **`id`**, `top`, `skipToken` |
| `directory_list_groups` |  | Lists groups in the Microsoft Entra ID tenant, 25 per page by default (max 999). | `top`, `filter`, `select`, `skipToken` |
| `directory_get_group` |  | Returns one group by object id, with its description, mail nickname, type flags, visibility, classification, lifecycle dates, dynamic membership rule and on-premises sync fields. | **`id`** |
| `directory_list_group_members` |  | Lists the direct members of a group, 25 per page by default (max 999). | **`id`**, `top`, `skipToken` |
| `directory_list_user_groups` |  | Returns the ids of every group the user belongs to, including memberships inherited through nested groups. | **`id`** |

</details>

<details>
<summary><b><code>intune</code></b> — Intune: 6 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `intune_list_managed_devices` |  | Lists Intune-managed devices in the tenant with id, name, operating system and version, compliance state, last sync time, primary user, manufacturer and model, plus count and nextLink when more pages exist. | `top`, `filter`, `select` |
| `intune_get_managed_device` |  | Reads one Intune-managed device in detail: hardware identity, primary user, enrollment and management state, encryption and supervision flags, Exchange access state, storage and memory in GB, and the last sync time. | **`deviceId`** |
| `intune_list_device_apps` |  | Lists the applications Intune has detected on one managed device, returning display name, version, size in MB and the tenant-wide deviceCount for each, with count and nextLink. | **`deviceId`**, `top` |
| `intune_list_compliance_policies` |  | Lists the tenant's Intune device compliance policies with id, display name, platform-specific policyType (from the @odata.type discriminator, e.g. | `top` |
| `intune_get_device_summary` |  | Returns the tenant-wide Intune device counts in one call: total enrolled devices, MDM-enrolled and dual-enrolled counts, a per-operating-system breakdown, and the Exchange access state breakdown. | none |
| `intune_list_noncompliant` |  | Lists managed devices whose compliance state is exactly 'noncompliant', with the usual device fields plus the compliance grace period expiry, count and nextLink. | `top` |

</details>

<details>
<summary><b><code>generic</code></b> — Generic: 6 tools, none of them writes</summary>

Always enabled and not selectable. Signing in lives here because it is needed before anything else works, and `graph_request` is the escape hatch for endpoints no purpose-built tool covers.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `auth_begin_login` |  | Opens the system browser for OAuth 2.0 authorization code + PKCE sign-in and returns immediately. Reuses an in-progress attempt instead of opening another window. Pass `scopes` to grow consent incrementally, `mode: "full"` to consent up front to everything grantable. No token is ever returned. | `scopes`, `mode` |
| `auth_login_status` |  | Reports where a browser sign-in has got to: pending, signedIn, failed or none. Never returns a token. | none |
| `session_status` |  | Unified session and consent probe: cached account plus whether the stored session still mints tokens and live GET /me reachability. Call once first. | none |
| `graph_request` | write | Sends an arbitrary request to Microsoft Graph and returns the raw JSON response. GET is always available; the write methods need `--allow-generic-write` and are refused under `--read-only`. | **`path`**, `method`, `version`, `query`, `body`, `maxPages` |
| `graph_schema` |  | Fetches one item from a Graph path ($top=1) and reports the property names it carries with their JSON types, plus the @odata.context that names the resource type. | **`entityPath`**, `version` |
| `graph_list_permissions` |  | Reports how this server is configured: which tool groups are enabled, the delegated Microsoft Graph scopes each one uses, which of them need a tenant administrator to consent, whether write tools are suppressed, and the signed-in account. | none |

</details>

## Document parsing

`files_parse_document` extracts `.docx`, `.xlsx`, `.pptx` and `.pdf` from a sharing link, a drive address, or a mail attachment, as unit rows (sections/sheets/slides/pages) plus a manifest with a cursor for paging. The manifest also carries a whole-document `titles` outline (every unit's id and title), so surveying a document needs no text fetch. Tables render as Markdown inside unit text (single-record blocks transpose to Field/Value form); `.docx` titles carry their heading path, headers/footers append as marked units, tracked insertions tag inline (`[+...+]`) with deletions collected into a marked unit, and comments append at the end; `.xlsx` cells render as seen (dates, currency, percent, 1904 system), hidden sheets/rows skip with a warning, and merged ranges fill from the top-left value; `.pptx` resolves cached chart values, SmartArt text and embedded workbooks in place, dedupes repeated footer boilerplate, and skips note placeholders. Caps: 16 MiB per download, 50 PDF pages, 100,000 spreadsheet rows across all sheets (100 columns each); overruns stop with a warning, not a silent cut. Non-goals: legacy `.doc`/`.ppt`/`.xls` are refused rather than converted, scanned or image-only PDFs are refused rather than OCR'd, password-protected files are refused rather than unlocked, images are not extracted, and chart/SmartArt graphics without cached values warn rather than guess.

## TOON output

`--output toon` serialises tool results as Token-Oriented Object Notation instead of indented JSON: uniform arrays become one header plus bare data rows, which spends far less of the context budget on repeated keys. The encoder implements the subset tool results need (objects, nested arrays, scalar quoting with the spec's escape set) and survives hostile values the same way JSON does — cycles, BigInts and Errors never throw. It deliberately omits keyed-tabular form, non-comma delimiters and tab indentation.

## Presets

| Preset | Groups |
| --- | --- |
| `personal` | me, mail, calendar, files, todo, contacts, search |
| `work` | personal plus chat, planner, sharepoint |
| `admin` | directory, intune, teams |
| `all` | every group |

`personal` is the default when neither `--groups` nor `--preset` is given. `--groups` and `--preset` can be combined; their union is taken.
