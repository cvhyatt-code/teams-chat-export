# Export-TeamsChat

Export a Microsoft Teams 1:1 or group chat to searchable text, JSON, and images —
including chats with people who have left and now show as **Unknown User**.

No Purview. No eDiscovery. No admin rights in most tenants. No paid tool.

## Why this exists

Search "export Teams chat" and every answer assumes you're an administrator
exporting *somebody else's* conversation, so they all point at Purview
eDiscovery or the Teams Export APIs. If the chat is **yours**, that's the wrong
problem. Your copy of the thread lives in your own chat store, and Microsoft
Graph will hand it to you with an ordinary user token.

The departed colleague doesn't matter. Their account can be fully deleted from
the directory and your side of the conversation is still intact — including,
usefully, their display name on every message they sent.

## Quick start

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
.\Export-TeamsChat.ps1
```

Run it with no arguments and it asks you for what it needs:

```
  Export-TeamsChat
  Pulls one Teams conversation into searchable text, JSON and images.

  In Teams: right-click the chat in the left rail -> Copy link.
  Or open the chat at teams.cloud.microsoft and copy the address bar.

  Paste the chat link: _
```

It signs you in, shows you who's actually in the chat and when it last had
traffic, waits for you to confirm, then runs — and prints the equivalent
one-liner so your next run can skip the questions:

```powershell
.\Export-TeamsChat.ps1 -FromLink "https://teams.cloud.microsoft/l/chat/19:...@unq.gbl.spaces/conversations?context=..." -Html
```

A browser opens and you consent to `Chat.Read` for yourself. Whether that's
allowed depends on one tenant setting, and you'll know in thirty seconds:

- **Consent screen with an Accept button** — you're done, nothing else to set up.
- **"Approval required" / `AADSTS65001`** — your tenant restricts which apps users
  may consent to. See [Locked-down tenants](#locked-down-tenants).

`Chat.Read` shows **Admin consent required: No** on the permissions blade, which
means the permission isn't inherently admin-gated. It does *not* mean your tenant
lets you consent to it. Tenants set user consent to one of: all apps, verified
publishers with low-impact permissions only, or blocked. Only the first lets this
run with no setup; `Chat.Read` is not a low-impact sign-in scope, so the middle
setting will stop you too.

Either way, this only ever reads **your own** chats. `Chat.Read` is delegated —
it returns the signed-in user's conversations and nothing else. There is no
configuration of this script that reads someone else's chat.

## What you get

```
TeamsExport\chat_20260914-151200\
  transcript.html                           one file, everything in it (-Html)
  chunks\
    part-01_2019-03-14_to_2021-08-02.md    self-contained, ~45k tokens each
    part-02_2021-08-03_to_2024-04-15.md
  images\
    img-0001.png                            every pasted screenshot, in order
    img-0042.png
    index.md                                what each screenshot is and where it sat
  transcript.jsonl                          one JSON object per message
  manifest.md                               roster, counts, image index
```

Two artifacts, two jobs.

**`transcript.html`** (`-Html`) is the one for humans: a single self-contained
file with every screenshot embedded inline, a live search box, click-to-zoom on
images, dark mode, and print styles. Open it in a browser, Ctrl+F to find things,
Ctrl+P if you need a PDF. Nothing to unzip and no folder of loose images to open
one at a time. Email it, archive it, hand it to legal.

If you want the PDF non-interactively:

```powershell
msedge --headless --disable-gpu --print-to-pdf="transcript.pdf" "transcript.html"
```

PDF is deliberately not the default. Pagination chops conversations mid-thread,
search is worse, and it's a one-way trip. HTML gives you the PDF whenever you
want it and stays useful in the meantime.

**The markdown chunks** are for feeding to a model. A years-long chat as one HTML
file can run to hundreds of megabytes, which no model will take. The chunks are
sized to upload.

Each chunk carries its own header so it makes sense uploaded alone, consecutive
messages from one speaker collapse into a single block, and Teams reply-quotes
are reduced to a short marker instead of inlining a full copy of the quoted
message — which otherwise doubles your token count for no information.

**Images are first class.** Pasted screenshots appear as
`[IMAGE: images/img-0042.png]` at the exact point they were sent. This matters
more than it sounds: in a long working chat, a startling amount of the real
content — forwarded threads, letters, error messages, whiteboard photos — exists
only as pixels. A transcript that renders those as `[image]` throws away the part
you were trying to keep.

## Getting screenshots into an analysis

The chunks reference screenshots as `[IMAGE: images/img-0042.png]` rather than
containing them, so a chunk alone tells a model that a screenshot existed but not
what was in it. In a working chat that can be a lot of the actual content —
forwarded threads, letters, error messages, whiteboard photos.

The filenames are the join key. Upload the chunk and the relevant PNGs together
and add one line to your prompt:

> The attached PNG files are screenshots from this conversation. Their filenames
> match the `[IMAGE: images/img-NNNN.png]` markers in the transcript — use the
> filename to place each one at the right point.

Don't embed images as base64 data URIs in the markdown. It's valid markdown, but
the upload path won't render it and you've just handed the model a few hundred
kilobytes of noise per image.

For a chat with a lot of screenshots, uploading all of them is waste. Use
`images\index.md` instead: it lists every screenshot with its date, sender, and
the conversation immediately before and after it. Hand a model that file plus the
PNGs and it can place each one in context without the transcript at all — which
also makes it easy to decide which twenty are worth reading properly.

See `ANALYSIS.md` for the two-pass extraction approach this feeds into.

## Finding the chat

Easiest is `-FromLink`. If you can't get a link — the chat is buried, or you're
looking for one among thousands — the script can hunt:

```powershell
# Inventory every chat to TeamsExport\chats.csv, showing 1:1s whose
# other participant no longer exists in the directory
.\Export-TeamsChat.ps1 -List -OrphanedOnly

# Then resolve who is ACTUALLY in each 1:1 by reading sender names off the
# messages, which survive account deletion
.\Export-TeamsChat.ps1 -Identify -Filter2 "<surname>"

# Confirm before committing to a full pull
.\Export-TeamsChat.ps1 -Peek -ChatId "19:..."
```

Two things worth knowing, both of which will waste your time otherwise:

- **`lastUpdatedDateTime` is not the last message time.** It's when the chat was
  renamed or its membership changed. Sorting by it to find a thread that went
  quiet in 2024 does not work.
- **Deleted users vanish from the member list entirely.** Graph doesn't return
  them as "Unknown User" — it returns a 1:1 chat with only you in it. That's the
  fingerprint to look for, and it's why `-Identify` reads names off messages
  instead of off membership.

## Options

| Flag | Effect |
|---|---|
| `-FromLink <url>` | Take the chat id from a Teams deep link |
| `-ChatId <id>` | Or specify it directly |
| `-Peek` | Print the last 20 messages and stop |
| `-List` | Inventory all chats to `chats.csv` |
| `-OrphanedOnly` | With `-List`: only chats whose counterpart is gone |
| `-Identify` | Resolve who's in each 1:1 from message senders |
| `-Html` | Also write one self-contained `transcript.html` with images embedded |
| `-HtmlLinkImages` | Reference images instead of embedding — smaller file, must stay beside `images\` |
| `-NoImages` | Skip the image download (much faster) |
| `-NameMap @{'<guid>'='Name'}` | Name anyone Graph returns without a display name |
| `-MaxCharsPerChunk <n>` | Chunk size, default 180000 (~45k tokens) |
| `-KeepSystemMessages` | Keep joins, renames, call events |
| `-KeepReactions` | Keep reaction metadata |
| `-KeepFullUrls` | Keep full URLs instead of just the hostname |
| `-ClientId` / `-TenantId` | Use your own app registration |

## Locked-down tenants

If you get **AADSTS65001** or an "Approval required" screen, your tenant has
turned off user consent to applications.

Do **not** submit the approval request for "Microsoft Graph Command Line Tools."
That app is shared by everyone with the Graph module installed; consenting
`Chat.Read` on it is a permanent tenant-wide grant on an app you don't control,
to solve a one-time problem for one person.

Register your own instead. Requires **Cloud Application Administrator** — which
is enough because `Chat.Read` is a *delegated* permission. (Application
permissions on Microsoft Graph would need Privileged Role Administrator; you're
not requesting one.)

1. **Entra portal → App registrations → New registration.** Single tenant.
2. **Authentication → Add a platform → Mobile and desktop applications.** Tick
   `https://login.microsoftonline.com/common/oauth2/nativeclient`, add
   `http://localhost`, and add
   `ms-appx-web://Microsoft.AAD.BrokerPlugin/<your-client-id>`.
   That third one is the Windows account broker, which recent Graph SDK versions
   use by default. Leave it out and you get a redirect-URI mismatch that reads
   like a permissions failure.
3. **API permissions → Microsoft Graph → Delegated → `Chat.Read`.** Exactly that
   one. Not `Chat.Read.All`, not `Chat.ReadWrite` — if you grab a `.All` variant
   you've built the thing you were trying to avoid.
4. **Grant admin consent.**
5. **Enterprise applications → your app → Properties → Assignment required: Yes.**
6. **Users and groups → assign only the account whose chats you're exporting.**
   Not an admin account — the chats live in the normal mailbox, and assigning the
   wrong one produces AADSTS50105.

Then:

```powershell
Disconnect-MgGraph
.\Export-TeamsChat.ps1 -FromLink "<link>" -ClientId "<app id>" -TenantId "contoso.com"
```

Delete the registration when you're done. A scoped, single-user, time-boxed grant
with a clean audit trail is a much easier conversation than a standing tenant-wide
one.

The script recognizes the common Entra failures and prints what each actually
means, which is rarely what the error text suggests.

## Limits

- **Attachments come through as links, not files.** If the sender's OneDrive was
  purged when they left, those links are dead. Check this early if the material
  you care about is in attachments rather than inline.
- **Images need a model to be searchable.** The script preserves and positions
  them; turning the pixels back into text is a separate step.
- **Large mailboxes are slow to enumerate.** `-List` on several thousand chats
  takes a few minutes. `-FromLink` skips it entirely.
- **Throttling is normal.** Image-heavy chats hit HTTP 429 constantly. The script
  backs off and continues; let it run.

## Requirements

PowerShell 5.1 or 7.x, `Microsoft.Graph.Authentication`, and a machine with a
browser. Device code flow is deliberately not used — it's blocked by Conditional
Access in many tenants, and the resulting error looks like something else.

## License

MIT.
