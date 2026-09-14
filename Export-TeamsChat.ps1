<#
.SYNOPSIS
    Exports a Microsoft Teams chat into AI-analysis-ready chunks, with inline
    images preserved and referenced by filename. Delegated Graph auth only --
    no admin rights, no Purview, no paid tool.

.DESCRIPTION
    Works on YOUR copy of the chat. The other participant showing as
    "Unknown User" in the Teams header does not matter: the thread lives in your
    own chat store, and per-message sender display names usually still resolve.

    Output:
      chunks\part-NN_<daterange>.md   self-contained markdown, ~45k tokens each
      images\img-NNNN.png             every pasted screenshot, in order
      transcript.jsonl                one message per line, full fidelity
      manifest.md                     roster, counts, and an image index

    Images are downloaded by default and referenced inline as
    [IMAGE: images/img-0042.png] at the exact point they appeared, so an LLM
    reading a chunk knows where each screenshot belongs.

.PARAMETER NameMap
    Fallback for senders Graph returns with no display name.
        -NameMap @{ '<guid>' = 'Jane Smith' }

.EXAMPLE
    .\Export-TeamsChat.ps1 -List -Filter "Unknown"
    .\Export-TeamsChat.ps1 -ChatId "19:xxx@unq.gbl.spaces"
#>

[CmdletBinding(DefaultParameterSetName = 'List')]
param(
    [Parameter(ParameterSetName = 'List')]
    [switch]$List,

    [Parameter(ParameterSetName = 'List')]
    [string]$Filter,

    # Show only 1:1 chats whose other participant no longer exists in the
    # directory. Graph drops deleted users from the member list entirely, so
    # these show up as a oneOnOne chat with only you in it.
    [Parameter(ParameterSetName = 'List')]
    [switch]$OrphanedOnly,

    # Walk the chats.csv produced by -List, sample each 1:1 chat, and record who
    # the other party actually is -- resolved from message sender names, which
    # survive account deletion even when directory membership does not.
    [Parameter(ParameterSetName = 'Identify', Mandatory = $true)]
    [switch]$Identify,

    [Parameter(ParameterSetName = 'Identify')]
    [string]$FromCsv,

    [Parameter(ParameterSetName = 'Identify')]
    [string]$Filter2,

    [Parameter(ParameterSetName = 'Export')]
    [string]$ChatId,

    # Paste a Teams deep link instead of hunting for a chat id:
    # right-click the chat in Teams -> Copy link, or grab the URL from the web
    # client. Looks like https://teams.cloud.microsoft/l/chat/19:...@unq.gbl.spaces/conversations?context=...
    [Parameter(ParameterSetName = 'Export')]
    [string]$FromLink,

    # Print the first and last few messages of a chat instead of exporting it,
    # so you can identify a thread before committing to a full pull.
    [Parameter(ParameterSetName = 'Export')]
    [switch]$Peek,

    [string]$OutDir = (Join-Path (Get-Location) 'TeamsExport'),

    # Own single-tenant app registration, if the tenant blocks user consent.
    [string]$ClientId,
    [string]$TenantId,

    [Parameter(ParameterSetName = 'Export')]
    [hashtable]$NameMap = @{},

    [Parameter(ParameterSetName = 'Export')]
    [int]$MaxCharsPerChunk = 180000,

    [Parameter(ParameterSetName = 'Export')]
    [switch]$NoImages,

    # Also emit one self-contained transcript.html with every screenshot embedded
    # inline, a live search box, and print styles. Open it in a browser; Ctrl+P
    # for a PDF. Nothing to unzip, no folder of loose images.
    [Parameter(ParameterSetName = 'Export')]
    [switch]$Html,

    # Reference image files instead of embedding them. Produces a much smaller
    # HTML file that must stay next to the images folder. Use when embedding
    # would push the single file past a few hundred MB.
    [Parameter(ParameterSetName = 'Export')]
    [switch]$HtmlLinkImages,

    [Parameter(ParameterSetName = 'Export')]
    [switch]$KeepSystemMessages,

    [Parameter(ParameterSetName = 'Export')]
    [switch]$KeepReactions,

    [Parameter(ParameterSetName = 'Export')]
    [switch]$KeepFullUrls
)

$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------
# Teams deep link -> chat id
# --------------------------------------------------------------------------
function ConvertFrom-TeamsLink {
    param([Parameter(Mandatory)][string]$Link)

    # Channel links are a different animal (they need teamId/channelId and a
    # different endpoint). Fail loudly rather than half-working.
    if ($Link -match '/l/channel/') {
        throw "That's a channel link, not a chat link. This script exports 1:1 and group chats. Channel messages live under /teams/{id}/channels/{id}/messages."
    }

    if ($Link -notmatch '/l/chat/([^/?]+)') {
        throw "Couldn't find a chat id in that link. Expected something like https://teams.cloud.microsoft/l/chat/19:...@unq.gbl.spaces/conversations?context=..."
    }

    $id = [uri]::UnescapeDataString($Matches[1])
    if ($id -notmatch '^19:') {
        throw "Extracted '$id', which doesn't look like a Teams chat id (should start with '19:')."
    }
    return $id
}

# --------------------------------------------------------------------------
# Translate Entra sign-in failures into what to actually do about them.
# Every one of these cost a round trip to discover the hard way.
# --------------------------------------------------------------------------
function Show-AuthGuidance {
    param([Parameter(Mandatory)][string]$Message)

    $hint = switch -Regex ($Message) {
        'AADSTS65001|need admin approval|Approval required' {
@"
Your tenant does not allow users to consent to applications, so the stock
"Microsoft Graph Command Line Tools" app can't be used here.

Do NOT submit the approval request for that app -- it's shared by everyone with
the Graph module installed, and consenting Chat.Read on it is a tenant-wide grant
on an app you don't control.

Register your own single-user app instead. See "Locked-down tenants" in the
README; it's six portal steps and takes about five minutes.
"@
        }
        'AADSTS700016' {
@"
The client id wasn't found in this tenant. Either it's mistyped (copy it with the
portal's copy button -- transcribing GUIDs by eye is how this usually happens),
or you're signing in to a different tenant than the one the app lives in, or the
registration is less than a few minutes old and hasn't replicated yet.
"@
        }
        'AADSTS50011' {
@"
Redirect URI mismatch. The error text names the exact URI that was requested --
add that one to your app registration under Authentication -> Mobile and desktop
applications.

If it starts with ms-appx-web://Microsoft.AAD.BrokerPlugin/, that's the Windows
account broker, which recent Graph SDK versions use by default. The URI is
ms-appx-web://Microsoft.AAD.BrokerPlugin/<your client id>. Nobody guesses this
one; it just has to be added.
"@
        }
        'AADSTS50105' {
@"
The signed-in account isn't assigned to the app. If you set "Assignment required"
to Yes, go to Enterprise applications -> your app -> Users and groups and assign
the account whose chats you're exporting.

Watch out for admin vs. normal accounts here: the chats live in the normal
account, so that's the one that needs the assignment.
"@
        }
        'AADSTS7000218' {
@"
The app is being treated as a confidential client. In the app registration, go to
Authentication -> Advanced settings and set "Allow public client flows" to Yes.
"@
        }
        'AADSTS50005|AADSTS50076|Conditional Access|device code' {
@"
Conditional Access blocked this sign-in. Device code flow is blocked in many
tenants -- run this on a machine with a browser and don't pass
-UseDeviceAuthentication.
"@
        }
        default { $null }
    }

    if ($hint) {
        Write-Host "`n--- What this actually means ---" -ForegroundColor Yellow
        Write-Host $hint -ForegroundColor Yellow
        Write-Host "--------------------------------`n" -ForegroundColor Yellow
    }
}

# --------------------------------------------------------------------------
# Interactive wizard -- what you get when the script is run with no arguments.
# Every flag still works non-interactively; this just means a first-time user
# doesn't have to read the help to get anywhere.
# --------------------------------------------------------------------------
$script:Interactive = $false
$script:EquivalentCommand = ''

if ($PSBoundParameters.Count -eq 0) {
    $script:Interactive = $true

    Write-Host ""
    Write-Host "  Export-TeamsChat" -ForegroundColor Cyan
    Write-Host "  Pulls one Teams conversation into searchable text, JSON and images." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  In Teams: right-click the chat in the left rail -> Copy link." -ForegroundColor Gray
    Write-Host "  Or open the chat at teams.cloud.microsoft and copy the address bar." -ForegroundColor Gray
    Write-Host ""

    $answer = (Read-Host "  Paste the chat link").Trim().Trim('"')

    if (-not $answer) {
        Write-Host ""
        Write-Host "  No link given. If you can't find the chat -- for example the other person" -ForegroundColor Yellow
        Write-Host "  has left and it shows as Unknown User -- these will hunt for it:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "    .\Export-TeamsChat.ps1 -List -OrphanedOnly" -ForegroundColor White
        Write-Host "    .\Export-TeamsChat.ps1 -Identify -Filter2 `"<surname>`"" -ForegroundColor White
        Write-Host ""
        return
    }

    if ($answer -match '^https?://') { $FromLink = $answer }
    elseif ($answer -match '^19:')   { $ChatId   = $answer }
    else {
        Write-Error "That doesn't look like a Teams chat link or a chat id (ids start with '19:')."
        return
    }

    Write-Host ""
    Write-Host "  Most tenants let you sign in directly -- just press Enter." -ForegroundColor DarkGray
    Write-Host "  Only paste a client id if your tenant blocks user consent and you've" -ForegroundColor DarkGray
    Write-Host "  registered your own app (see the README)." -ForegroundColor DarkGray
    $cid = (Read-Host "  Client id [Enter to skip]").Trim()
    if ($cid) {
        $ClientId = $cid
        $tid = (Read-Host "  Tenant id or domain").Trim()
        if ($tid) { $TenantId = $tid }
    }

    Write-Host ""
    $ans = (Read-Host "  Download pasted screenshots? Slower, but they often hold the real content [Y/n]").Trim()
    if ($ans -match '^n') { $NoImages = $true }

    $ans = (Read-Host "  Build a single-file HTML transcript you can open and search? [Y/n]").Trim()
    if ($ans -notmatch '^n') { $Html = $true }

    # Show them the non-interactive form so the second run is one line.
    $parts = @('.\Export-TeamsChat.ps1')
    if ($FromLink) { $parts += "-FromLink `"$FromLink`"" } else { $parts += "-ChatId `"$ChatId`"" }
    if ($Html)     { $parts += '-Html' }
    if ($NoImages) { $parts += '-NoImages' }
    if ($ClientId) { $parts += "-ClientId `"$ClientId`"" }
    if ($TenantId) { $parts += "-TenantId `"$TenantId`"" }
    $script:EquivalentCommand = $parts -join ' '

    Write-Host ""
}

# --------------------------------------------------------------------------
# Connect
# --------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication (current user scope)..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

if (-not (Get-MgContext)) {
    # Interactive browser auth on purpose. Do NOT add -UseDeviceAuthentication:
    # device code flow is blocked by Conditional Access in some tenants and fails
    # with a policy error that looks like a permissions problem and isn't.
    $connect = @{ Scopes = @('Chat.Read'); NoWelcome = $true }
    if ($ClientId) { $connect.ClientId = $ClientId }
    if ($TenantId) { $connect.TenantId = $TenantId }

    if (-not $ClientId) {
        Write-Host "Signing in with the built-in Graph CLI app. In most tenants this just works --" -ForegroundColor DarkGray
        Write-Host "Chat.Read is user-consentable by default. If it doesn't, you'll get guidance." -ForegroundColor DarkGray
    }

    try { Connect-MgGraph @connect }
    catch {
        Write-Host "`nSign-in failed:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Show-AuthGuidance -Message $_.Exception.Message
        throw
    }
}
$ctx = Get-MgContext
Write-Host "Signed in as $($ctx.Account) (tenant $($ctx.TenantId))" -ForegroundColor Green

# --------------------------------------------------------------------------
# Graph plumbing
# --------------------------------------------------------------------------
function Invoke-GraphGet {
    param([Parameter(Mandatory)][string]$Uri, [switch]$Raw)
    $attempt = 0
    while ($true) {
        try {
            if ($Raw) { return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType HttpResponseMessage }
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        }
        catch {
            $status = $_.Exception.Response.StatusCode.value__
            if ($status -eq 429 -or $status -eq 503) {
                $attempt++
                if ($attempt -gt 6) { throw }
                $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Host "  throttled ($status) - sleeping $wait s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphAllPages {
    param([Parameter(Mandatory)][string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri; $page = 0
    while ($next) {
        $page++
        $resp = Invoke-GraphGet -Uri $next
        if ($resp.value) { $all.AddRange(@($resp.value)) }
        Write-Host ("  page {0,-4} total {1}" -f $page, $all.Count) -ForegroundColor DarkGray
        $next = $resp.'@odata.nextLink'
    }
    return $all
}

function Format-Members {
    param($Chat)
    $names = @()
    foreach ($m in @($Chat.members)) {
        $n = if ($m.displayName) { $m.displayName } else { 'Unknown User' }
        if ($m.email) { $n = "$n <$($m.email)>" }
        $names += $n
    }
    if (-not $names) { return '(no members returned - deleted account)' }
    return ($names -join ', ')
}

# --------------------------------------------------------------------------
# LIST mode
# --------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'List' -and -not $script:Interactive) {
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

    Write-Host "`nEnumerating chats (this takes a minute on a large mailbox)..." -ForegroundColor Cyan
    $chats = Get-GraphAllPages -Uri 'https://graph.microsoft.com/v1.0/me/chats?$expand=members&$top=50'

    $me = $ctx.Account
    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($c in $chats) {
        $others = @($c.members | Where-Object {
            $_.email -and ($_.email -ne $me)
        })
        $unknownMembers = @($c.members | Where-Object { -not $_.displayName })

        # A 1:1 with nobody but you left in it means the other party was deleted
        # from the directory. Same for any chat returning zero members.
        $orphaned = ($c.chatType -eq 'oneOnOne' -and $others.Count -eq 0) -or
                    (@($c.members).Count -eq 0) -or
                    ($unknownMembers.Count -gt 0)

        $rows.Add([pscustomobject]@{
            Index       = $rows.Count + 1
            ChatType    = $c.chatType
            Topic       = $(if ($c.topic) { $c.topic } else { '(1:1)' })
            Orphaned    = $orphaned
            MemberCount = $c.memberCount
            Members     = (Format-Members -Chat $c)
            LastUpdated = $c.lastUpdatedDateTime
            ChatId      = $c.id
        })
    }

    $csv = Join-Path $OutDir 'chats.csv'
    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host "`nWrote full chat inventory to $csv ($($rows.Count) chats)" -ForegroundColor Green

    $show = $rows
    if ($OrphanedOnly) { $show = $rows | Where-Object { $_.Orphaned } }
    if ($Filter) {
        $show = $show | Where-Object { $_.Members -match [regex]::Escape($Filter) -or $_.Topic -match [regex]::Escape($Filter) }
    }

    Write-Host ""
    foreach ($r in $show) {
        $flag = if ($r.Orphaned) { '  <-- deleted participant' } else { '' }
        Write-Host ("[{0}] {1}  {2}{3}" -f $r.Index, $r.ChatType.PadRight(10), $r.Topic, $flag) -ForegroundColor White
        Write-Host ("     with : {0}" -f $r.Members) -ForegroundColor Gray
        Write-Host ("     id   : {0}" -f $r.ChatId) -ForegroundColor DarkGray
        Write-Host ""
    }

    Write-Host ("{0} chat(s) shown of {1} total ({2} orphaned)." -f @($show).Count, $rows.Count, @($rows | Where-Object { $_.Orphaned }).Count) -ForegroundColor Cyan
    Write-Host "Next: .\export-chat.ps1 -ChatId `"<id>`" -Peek   to see who's in it before exporting.`n" -ForegroundColor Yellow
    return
}

# --------------------------------------------------------------------------
# IDENTIFY mode -- who is actually in each 1:1 chat
# --------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'Identify') {
    if (-not $FromCsv) { $FromCsv = Join-Path $OutDir 'chats.csv' }
    if (-not (Test-Path $FromCsv)) {
        Write-Error "No chat inventory at $FromCsv. Run with -List first."
        return
    }

    $me     = $ctx.Account
    $myName = ''
    try { $myName = (Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/me').displayName } catch { }
    if (-not $myName) {
        Write-Warning "Couldn't resolve your own display name, so the Counterpart column will also list you. Not fatal, just noisier."
    }

    $all = @(Import-Csv $FromCsv)
    $targets = @($all | Where-Object { $_.ChatType -eq 'oneOnOne' })
    Write-Host "`nIdentifying $($targets.Count) one-on-one chat(s) out of $($all.Count)..." -ForegroundColor Cyan
    Write-Host "One Graph call each. This takes a few minutes and will throttle; let it run.`n" -ForegroundColor DarkGray

    $out = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($row in $targets) {
        $n++
        if ($n % 20 -eq 0) { Write-Host "  $n / $($targets.Count)..." -ForegroundColor DarkGray }

        $eid = [uri]::EscapeDataString($row.ChatId)
        $counterpart = ''
        $lastMsg     = ''
        $msgSample   = ''
        try {
            $r = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/me/chats/$eid/messages?`$top=10"
            $msgs = @($r.value)
            if ($msgs.Count) {
                $lastMsg = ([datetime]$msgs[0].createdDateTime).ToLocalTime().ToString('yyyy-MM-dd')
                $others = @($msgs |
                    Where-Object { $_.from.user.displayName -and $_.from.user.displayName -ne $myName } |
                    ForEach-Object { $_.from.user.displayName } |
                    Select-Object -Unique)
                $counterpart = ($others -join '; ')
                $body = ($msgs[0].body.content -replace '<[^>]+>', ' ')
                $msgSample = ([System.Net.WebUtility]::HtmlDecode($body) -replace '\s+', ' ').Trim()
                if ($msgSample.Length -gt 80) { $msgSample = $msgSample.Substring(0, 80) }
            }
        }
        catch { $counterpart = "[error: $($_.Exception.Message)]" }

        if (-not $counterpart) { $counterpart = '(only me / no other sender in sample)' }

        $out.Add([pscustomobject]@{
            Counterpart = $counterpart
            LastMessage = $lastMsg
            Orphaned    = $row.Orphaned
            Sample      = $msgSample
            ChatId      = $row.ChatId
        })
    }

    $idCsv = Join-Path $OutDir 'identified.csv'
    $out | Sort-Object Counterpart | Export-Csv -Path $idCsv -NoTypeInformation -Encoding UTF8
    Write-Host "`nWrote $idCsv" -ForegroundColor Green

    if ($Filter2) {
        $hits = @($out | Where-Object { $_.Counterpart -match [regex]::Escape($Filter2) })
        Write-Host "`nMatches for '$Filter2':`n" -ForegroundColor Cyan
        foreach ($h in $hits) {
            Write-Host ("  {0}   last message {1}" -f $h.Counterpart, $h.LastMessage) -ForegroundColor White
            Write-Host ("  {0}" -f $h.ChatId) -ForegroundColor Yellow
            Write-Host ("  `"{0}...`"`n" -f $h.Sample) -ForegroundColor DarkGray
        }
        if (-not $hits) { Write-Host "  none`n" -ForegroundColor DarkGray }
    }
    else {
        Write-Host "`nSearch it with:" -ForegroundColor Yellow
        Write-Host "  Import-Csv `"$idCsv`" | Where-Object Counterpart -match '<surname>' | Format-List`n" -ForegroundColor Yellow
    }
    return
}

# --------------------------------------------------------------------------
# Sender resolution
# --------------------------------------------------------------------------
$script:SenderRoster = @{}

function Resolve-Sender {
    param($Message)
    $u = $Message.from.user
    $guid = $null; $name = $null
    if ($u) { $guid = $u.id; $name = $u.displayName }
    elseif ($Message.from.application.displayName) { $name = "[bot] $($Message.from.application.displayName)" }

    if ($guid -and $NameMap.ContainsKey($guid))      { $name = $NameMap[$guid] }
    elseif ($name -and $NameMap.ContainsKey($name))  { $name = $NameMap[$name] }
    elseif (-not $name -and $NameMap.ContainsKey('Unknown User')) { $name = $NameMap['Unknown User'] }

    if (-not $name) { $name = if ($guid) { 'Unknown User' } else { 'System' } }

    $key = if ($guid) { $guid } else { $name }
    if (-not $script:SenderRoster.ContainsKey($key)) {
        $script:SenderRoster[$key] = @{ Name = $name; Count = 0; Guid = $guid }
    }
    $script:SenderRoster[$key].Count++
    $script:SenderRoster[$key].Name = $name
    return $name
}

# --------------------------------------------------------------------------
# HTML -> clean text, with positional image references
# --------------------------------------------------------------------------
function ConvertFrom-TeamsHtml {
    param([string]$Html, [string[]]$ImageRefs = @())

    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $t = $Html

    # Teams reply-quotes inline a full copy of the quoted message. Collapse to a
    # short marker rather than deleting, so the reference survives without the
    # token cost of duplicating every quoted message.
    $t = [regex]::Replace($t, '(?is)<blockquote[^>]*itemtype="http://schema\.skype\.com/Reply"[^>]*>(.*?)</blockquote>', {
        param($m)
        $inner = ($m.Groups[1].Value -replace '<[^>]+>', ' ')
        $inner = [System.Net.WebUtility]::HtmlDecode($inner) -replace '\s+', ' '
        $inner = $inner.Trim()
        if ($inner.Length -gt 120) { $inner = $inner.Substring(0, 120) + '...' }
        "[in reply to: `"$inner`"] "
    })

    # substitute <img> tags positionally with the files we downloaded
    $script:imgCursor = 0
    $t = [regex]::Replace($t, '(?i)<img[^>]*>', {
        param($m)
        $ref = if ($script:imgCursor -lt $ImageRefs.Count) { $ImageRefs[$script:imgCursor] } else { $null }
        $script:imgCursor++
        if ($ref) { "`n[IMAGE: $ref]`n" } else { "`n[IMAGE: not downloaded]`n" }
    })

    $t = $t -replace '(?i)<br\s*/?>', "`n"
    $t = $t -replace '(?i)</(p|div|li|tr|h[1-6])>', "`n"
    $t = $t -replace '(?i)<li[^>]*>', '- '

    if (-not $KeepFullUrls) {
        $t = [regex]::Replace($t, '(?is)<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', {
            param($m)
            $href = $m.Groups[1].Value
            $text = ($m.Groups[2].Value -replace '<[^>]+>', '').Trim()
            $h = try { ([uri]$href).Host } catch { 'link' }
            if ($text -and $text -notmatch '^https?://') { "$text [link: $h]" } else { "[link: $h]" }
        })
    }

    $t = $t -replace '<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace '[ \t]+', ' '
    $t = $t -replace "\n{3,}", "`n`n"
    return $t.Trim()
}

# --------------------------------------------------------------------------
# EXPORT
# --------------------------------------------------------------------------
if ($FromLink) {
    $ChatId = ConvertFrom-TeamsLink -Link $FromLink
    Write-Host "Chat id from link: $ChatId" -ForegroundColor Green
}
if (-not $ChatId) {
    Write-Error "Give me either -ChatId or -FromLink (right-click the chat in Teams -> Copy link)."
    return
}

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$encId = [uri]::EscapeDataString($ChatId)

Write-Host "`nReading chat metadata..." -ForegroundColor Cyan
$chat = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId`?`$expand=members"
$participants = Format-Members -Chat $chat
Write-Host "With : $participants" -ForegroundColor Green

# In wizard mode, show what we're about to pull and let them back out. Membership
# alone is unreliable for departed people, so read the names off the messages.
if ($script:Interactive) {
    try {
        $probe = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId/messages?`$top=10"
        $pm = @($probe.value)
        if ($pm.Count) {
            $who = @($pm | Where-Object { $_.from.user.displayName } |
                     ForEach-Object { $_.from.user.displayName } | Select-Object -Unique)
            Write-Host ("Senders  : {0}" -f ($who -join ', ')) -ForegroundColor Green
            Write-Host ("Latest   : {0}" -f ([datetime]$pm[0].createdDateTime).ToLocalTime().ToString('yyyy-MM-dd')) -ForegroundColor Green
        }
    } catch { }

    Write-Host ""
    $go = (Read-Host "Export this conversation? [Y/n]").Trim()
    if ($go -match '^n') { Write-Host "Cancelled.`n"; return }
    Write-Host ""
}

# ---- peek: identify a thread without pulling all of it ----
if ($Peek) {
    Write-Host "`nMost recent 20 messages:`n" -ForegroundColor Cyan
    $sample = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId/messages?`$top=20"
    $names = @{}
    foreach ($m in @($sample.value)) {
        $who = if ($m.from.user.displayName) { $m.from.user.displayName }
               elseif ($m.from.user.id)      { "Unknown <$($m.from.user.id)>" }
               else                          { '[system]' }
        if ($who -ne '[system]') { $names[$who] = $true }
        $body = ($m.body.content -replace '<[^>]+>', ' ')
        $body = ([System.Net.WebUtility]::HtmlDecode($body) -replace '\s+', ' ').Trim()
        if ($body.Length -gt 110) { $body = $body.Substring(0, 110) + '...' }
        if (-not $body) { $body = '(no text - image or system event)' }
        Write-Host ("  {0}  {1}" -f ([datetime]$m.createdDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm'), $who) -ForegroundColor White
        Write-Host ("      {0}" -f $body) -ForegroundColor Gray
    }
    Write-Host "`nSenders in this sample: $($names.Keys -join ', ')" -ForegroundColor Green
    Write-Host "If that's the right thread, re-run without -Peek to export it.`n" -ForegroundColor Yellow
    return
}

Write-Host "`nPulling messages (50/page)..." -ForegroundColor Cyan
$raw = Get-GraphAllPages -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId/messages?`$top=50"
if (-not $raw -or $raw.Count -eq 0) { Write-Warning "No messages returned. Check the chat id."; return }

$ordered = $raw | Sort-Object { [datetime]$_.createdDateTime }

$stamp    = Get-Date -Format 'yyyyMMdd-HHmmss'
$runDir   = Join-Path $OutDir "chat_$stamp"
$chunkDir = Join-Path $runDir 'chunks'
$imgDir   = Join-Path $runDir 'images'
New-Item -ItemType Directory -Path $chunkDir -Force | Out-Null
if (-not $NoImages) { New-Item -ItemType Directory -Path $imgDir -Force | Out-Null }

# ---- pass 1: download images so we can reference them positionally ----
$imgIndex   = [System.Collections.Generic.List[object]]::new()
$imgCounter = 0
$imgByMsg   = @{}

if (-not $NoImages) {
    $withImages = @($ordered | Where-Object { $_.body.content -match '<img' })
    Write-Host "`nDownloading inline images from $($withImages.Count) message(s)..." -ForegroundColor Cyan
    foreach ($m in $withImages) {
        $refs = @()
        try {
            $hc = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId/messages/$($m.id)/hostedContents"
            foreach ($h in @($hc.value)) {
                $imgCounter++
                $file = "img-{0:D4}.png" -f $imgCounter
                $resp = Invoke-GraphGet -Raw -Uri "https://graph.microsoft.com/v1.0/me/chats/$encId/messages/$($m.id)/hostedContents/$($h.id)/`$value"
                $bytes = $resp.Content.ReadAsByteArrayAsync().Result
                [System.IO.File]::WriteAllBytes((Join-Path $imgDir $file), $bytes)
                $refs += "images/$file"
                $imgIndex.Add([pscustomobject]@{
                    File   = "images/$file"
                    When   = ([datetime]$m.createdDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
                    Sender = $(if ($m.from.user.displayName) { $m.from.user.displayName } else { 'Unknown User' })
                    Bytes  = $bytes.Length
                })
            }
        } catch {
            Write-Warning "Image pull failed for message $($m.id): $($_.Exception.Message)"
        }
        if ($refs.Count) { $imgByMsg[$m.id] = $refs }
        if ($imgCounter % 25 -eq 0 -and $imgCounter -gt 0) { Write-Host "  $imgCounter images..." -ForegroundColor DarkGray }
    }
    Write-Host "Downloaded $imgCounter image(s)" -ForegroundColor Green
}

# ---- pass 2: normalize ----
$records = [System.Collections.Generic.List[object]]::new()
$skipped = @{ System = 0; Deleted = 0; Empty = 0 }

foreach ($m in $ordered) {
    if (($m.messageType -ne 'message') -and -not $KeepSystemMessages) { $skipped.System++; continue }

    $sender = Resolve-Sender -Message $m
    $refs   = if ($imgByMsg.ContainsKey($m.id)) { $imgByMsg[$m.id] } else { @() }
    $text   = ConvertFrom-TeamsHtml -Html $m.body.content -ImageRefs $refs

    if ($m.deletedDateTime) { $text = '[message deleted]'; $skipped.Deleted++ }

    $atts = @()
    foreach ($a in @($m.attachments)) {
        $an = if ($a.name) { $a.name } else { $a.contentType }
        if ($an) { $atts += $an }
    }
    if ($atts.Count) { $text = ($text + "`n[ATTACHMENT: " + ($atts -join '; ') + ']').Trim() }

    if ($KeepReactions -and $m.reactions) {
        $r = ($m.reactions | ForEach-Object { $_.reactionType }) -join ','
        if ($r) { $text = "$text`n(reactions: $r)" }
    }

    if ([string]::IsNullOrWhiteSpace($text)) { $skipped.Empty++; continue }

    $records.Add([pscustomobject]@{
        id        = $m.id
        timestamp = ([datetime]$m.createdDateTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        localTime = ([datetime]$m.createdDateTime).ToLocalTime()
        sender    = $sender
        senderId  = $m.from.user.id
        images    = $refs
        text      = $text
    })
}

if ($records.Count -eq 0) { Write-Warning "Everything was filtered out. Try -KeepSystemMessages."; return }

# ---- jsonl ----
$jsonlPath = Join-Path $runDir 'transcript.jsonl'
$sw = [System.IO.StreamWriter]::new($jsonlPath, $false, [System.Text.UTF8Encoding]::new($false))
foreach ($r in $records) {
    $sw.WriteLine(($r | Select-Object id, timestamp, sender, senderId, images, text | ConvertTo-Json -Compress -Depth 5))
}
$sw.Close()
Write-Host "`nWrote transcript.jsonl ($($records.Count) messages)" -ForegroundColor Green

# ---- markdown chunks ----
$first = $records[0].localTime
$last  = $records[-1].localTime

function New-ChunkHeader {
    param([int]$Index, $From, $To)
    @"
# Teams chat transcript - part $Index

**Participants:** $participants
**This part covers:** $($From.ToString('yyyy-MM-dd')) to $($To.ToString('yyyy-MM-dd'))
**Full conversation spans:** $($first.ToString('yyyy-MM-dd')) to $($last.ToString('yyyy-MM-dd'))
**Exported:** $(Get-Date -Format 'yyyy-MM-dd') by $($ctx.Account)

Notes for whoever reads this: timestamps are local time. Pasted screenshots appear
as ``[IMAGE: images/img-NNNN.png]`` at the point they were sent -- the image files
sit alongside this transcript and often carry content that exists nowhere in the
text. Reply-quotes are collapsed to a short marker. Join/leave events and
reactions were removed.

---

"@
}

$chunkIndex = 1
$buf = [System.Text.StringBuilder]::new()
$chunkStart = $records[0].localTime
$chunkFiles = @()
$lastDay = ''; $prevSender = ''; $prevTime = [datetime]::MinValue

function Save-Chunk {
    param($Index, $Start, $End, $Body)
    $name = "part-{0:D2}_{1}_to_{2}.md" -f $Index, $Start.ToString('yyyy-MM-dd'), $End.ToString('yyyy-MM-dd')
    ((New-ChunkHeader -Index $Index -From $Start -To $End) + $Body) |
        Out-File (Join-Path $chunkDir $name) -Encoding utf8
    return $name
}

for ($i = 0; $i -lt $records.Count; $i++) {
    $r = $records[$i]
    $day = $r.localTime.ToString('yyyy-MM-dd (dddd)')

    $wouldExceed = ($buf.Length -gt $MaxCharsPerChunk)
    $atDayBreak  = ($i -gt 0) -and ($r.localTime.Date -ne $records[$i-1].localTime.Date)
    if ($wouldExceed -and $atDayBreak) {
        $chunkFiles += Save-Chunk -Index $chunkIndex -Start $chunkStart -End $records[$i-1].localTime -Body $buf.ToString()
        $chunkIndex++
        $buf = [System.Text.StringBuilder]::new()
        $chunkStart = $r.localTime
        $lastDay = ''; $prevSender = ''
    }

    $block = ''
    if ($day -ne $lastDay) { $block += "`n## $day`n`n"; $lastDay = $day; $prevSender = '' }

    $sameRun = ($r.sender -eq $prevSender) -and (($r.localTime - $prevTime).TotalMinutes -lt 5)
    if ($sameRun) { $block += "$($r.text)`n`n" }
    else          { $block += "**$($r.sender)** _$($r.localTime.ToString('HH:mm'))_`n$($r.text)`n`n" }

    $prevSender = $r.sender
    $prevTime   = $r.localTime
    [void]$buf.Append($block)
}
if ($buf.Length -gt 0) { $chunkFiles += Save-Chunk -Index $chunkIndex -Start $chunkStart -End $last -Body $buf.ToString() }
Write-Host "Wrote $($chunkFiles.Count) chunk(s) to $chunkDir" -ForegroundColor Green

# ---- images/index.md : join table between PNG files and the conversation ----
# Lets you hand a model a pile of screenshots plus this one file and have it
# know exactly where each sat and what was being discussed around it.
if (-not $NoImages -and $imgCounter -gt 0) {

    function Get-Snippet {
        param($Record, [int]$Max = 220)
        if (-not $Record) { return '' }
        $t = ($Record.text -replace '\[IMAGE: [^\]]+\]', '[screenshot]') -replace '\s+', ' '
        $t = $t.Trim()
        if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max).TrimEnd() + '...' }
        return $t
    }

    $ix = [System.Text.StringBuilder]::new()
    [void]$ix.AppendLine("# Screenshot index")
    [void]$ix.AppendLine()
    [void]$ix.AppendLine("$imgCounter screenshot(s) from the conversation between $participants.")
    [void]$ix.AppendLine()
    [void]$ix.AppendLine("Each entry gives the file, when it was sent, who sent it, and the")
    [void]$ix.AppendLine("conversation immediately before and after it. Filenames match the")
    [void]$ix.AppendLine("``[IMAGE: images/img-NNNN.png]`` markers in the transcript chunks.")
    [void]$ix.AppendLine()
    [void]$ix.AppendLine("Upload this file alongside the PNGs and a model can place each screenshot")
    [void]$ix.AppendLine("in context without needing the full transcript.")
    [void]$ix.AppendLine()
    [void]$ix.AppendLine('---')
    [void]$ix.AppendLine()

    for ($i = 0; $i -lt $records.Count; $i++) {
        $r = $records[$i]
        # Guard against null/blank entries rather than trusting the array shape.
        $refs = @(@($r.images) | Where-Object { $_ -and "$_".Trim() })
        if ($refs.Count -eq 0) { continue }

        $prev = if ($i -gt 0) { $records[$i - 1] } else { $null }
        $next = if ($i -lt $records.Count - 1) { $records[$i + 1] } else { $null }
        for ($k = 0; $k -lt $refs.Count; $k++) {
            $file = Split-Path $refs[$k] -Leaf
            [void]$ix.AppendLine("## $file")
            [void]$ix.AppendLine()
            [void]$ix.AppendLine("- **Sent:** $($r.localTime.ToString('yyyy-MM-dd HH:mm')) ($($r.localTime.ToString('dddd')))")
            [void]$ix.AppendLine("- **By:** $($r.sender)")
            if ($refs.Count -gt 1) {
                [void]$ix.AppendLine("- **Position:** image $($k + 1) of $($refs.Count) in the same message")
            }
            [void]$ix.AppendLine()
            if ($prev) {
                [void]$ix.AppendLine("**Before** -- $($prev.sender) at $($prev.localTime.ToString('HH:mm')):")
                [void]$ix.AppendLine("> $(Get-Snippet -Record $prev)")
                [void]$ix.AppendLine()
            }
            [void]$ix.AppendLine("**Message containing it** -- $($r.sender):")
            $own = Get-Snippet -Record $r
            if (-not $own -or $own -eq '[screenshot]') { $own = '(no text -- the screenshot was the whole message)' }
            [void]$ix.AppendLine("> $own")
            [void]$ix.AppendLine()
            if ($next) {
                [void]$ix.AppendLine("**After** -- $($next.sender) at $($next.localTime.ToString('HH:mm')):")
                [void]$ix.AppendLine("> $(Get-Snippet -Record $next)")
                [void]$ix.AppendLine()
            }
            [void]$ix.AppendLine('---')
            [void]$ix.AppendLine()
        }
    }

    $ixPath = Join-Path $imgDir 'index.md'
    [System.IO.File]::WriteAllText($ixPath, $ix.ToString(), [System.Text.UTF8Encoding]::new($false))
    Write-Host "Wrote images\index.md ($imgCounter entries)" -ForegroundColor Green
}

# ---- single self-contained HTML transcript ----
if ($Html) {
    Write-Host "`nBuilding HTML transcript..." -ForegroundColor Cyan

    function ConvertTo-HtmlText {
        param([string]$s)
        if (-not $s) { return '' }
        $s = $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
        # linkify bare urls
        $s = [regex]::Replace($s, '(https?://[^\s<]+)', '<a href="$1" rel="noopener">$1</a>')
        return ($s -replace "`n", '<br>')
    }

    $imgCache = @{}
    function Get-ImgTag {
        param([string]$RelPath)
        if ($HtmlLinkImages) {
            return "<img class=`"shot`" src=`"$RelPath`" loading=`"lazy`" alt=`"pasted screenshot`">"
        }
        if (-not $imgCache.ContainsKey($RelPath)) {
            $full = Join-Path $runDir $RelPath
            if (Test-Path $full) {
                $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($full))
                $imgCache[$RelPath] = "data:image/png;base64,$b64"
            } else {
                $imgCache[$RelPath] = $null
            }
        }
        $src = $imgCache[$RelPath]
        if (-not $src) { return "<div class=`"missing`">[image not downloaded: $RelPath]</div>" }
        return "<img class=`"shot`" src=`"$src`" loading=`"lazy`" alt=`"pasted screenshot`">"
    }

    $myName = 'me'
    try { $myName = (Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/me').displayName } catch { }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine(@"
<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Teams chat - $([System.Net.WebUtility]::HtmlEncode($participants))</title>
<style>
:root{--bg:#f5f5f5;--card:#fff;--ink:#1b1b1b;--muted:#616161;--mine:#e3ecfb;--line:#e0e0e0;--accent:#5b5fc7}
@media(prefers-color-scheme:dark){:root{--bg:#1f1f1f;--card:#2b2b2b;--ink:#ebebeb;--muted:#a0a0a0;--mine:#2d3f5f;--line:#3d3d3d}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
header{position:sticky;top:0;z-index:10;background:var(--card);border-bottom:1px solid var(--line);padding:12px 20px}
header h1{margin:0 0 2px;font-size:16px}
header .meta{color:var(--muted);font-size:13px}
#q{margin-top:8px;width:100%;max-width:420px;padding:7px 10px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);font-size:14px}
#count{color:var(--muted);font-size:12px;margin-left:8px}
main{max-width:900px;margin:0 auto;padding:20px}
.day{position:sticky;top:96px;text-align:center;margin:26px 0 14px;color:var(--muted);font-size:12px;font-weight:600;letter-spacing:.04em;text-transform:uppercase}
.day span{background:var(--bg);padding:0 12px}
.day:before{content:"";display:block;border-top:1px solid var(--line);position:relative;top:9px;z-index:-1}
.msg{background:var(--card);border-radius:10px;padding:10px 14px;margin:0 0 8px;max-width:78%;box-shadow:0 1px 2px rgba(0,0,0,.06)}
.msg.mine{background:var(--mine);margin-left:auto}
.who{font-weight:600;font-size:13px;margin-bottom:2px}
.when{color:var(--muted);font-weight:400;font-size:12px;margin-left:8px}
.body a{color:var(--accent)}
.shot{display:block;max-width:100%;margin:8px 0;border:1px solid var(--line);border-radius:6px;cursor:zoom-in}
.shot.zoom{position:fixed;inset:20px;max-width:none;width:calc(100% - 40px);height:calc(100% - 40px);object-fit:contain;background:var(--card);z-index:99;cursor:zoom-out}
.missing{color:var(--muted);font-style:italic;font-size:13px}
.hidden{display:none}
@media print{
 header{position:static}.day{position:static}
 body{background:#fff}.msg{box-shadow:none;border:1px solid #ddd;max-width:100%;page-break-inside:avoid}
 #q,#count{display:none}.shot{max-height:16cm}
}
</style></head><body>
<header>
<h1>$([System.Net.WebUtility]::HtmlEncode($participants))</h1>
<div class="meta">$($records.Count) messages &middot; $($first.ToString('d MMM yyyy')) to $($last.ToString('d MMM yyyy')) &middot; $imgCounter screenshots &middot; exported $(Get-Date -Format 'yyyy-MM-dd') by $($ctx.Account)</div>
<input id="q" type="search" placeholder="Search this conversation..." autocomplete="off"><span id="count"></span>
</header>
<main>
"@)

    $lastDay = ''
    foreach ($r in $records) {
        $day = $r.localTime.ToString('dddd, d MMMM yyyy')
        if ($day -ne $lastDay) {
            [void]$sb.AppendLine("<div class=`"day`"><span>$day</span></div>")
            $lastDay = $day
        }

        $mine = if ($r.sender -eq $myName) { ' mine' } else { '' }
        $body = ConvertTo-HtmlText -s $r.text

        # swap the [IMAGE: images/x.png] markers for real <img> elements
        $body = [regex]::Replace($body, '\[IMAGE: ([^\]]+)\]', {
            param($m) Get-ImgTag -RelPath $m.Groups[1].Value
        })
        $body = $body -replace '\[IMAGE: not downloaded\]', '<span class="missing">[image not downloaded]</span>'

        [void]$sb.AppendLine("<div class=`"msg$mine`"><div class=`"who`">$([System.Net.WebUtility]::HtmlEncode($r.sender))<span class=`"when`">$($r.localTime.ToString('HH:mm'))</span></div><div class=`"body`">$body</div></div>")
    }

    [void]$sb.AppendLine(@'
</main>
<script>
const q=document.getElementById('q'),c=document.getElementById('count'),
      msgs=[...document.querySelectorAll('.msg')],days=[...document.querySelectorAll('.day')];
q.addEventListener('input',()=>{
  const t=q.value.trim().toLowerCase();
  if(!t){msgs.forEach(m=>m.classList.remove('hidden'));days.forEach(d=>d.classList.remove('hidden'));c.textContent='';return;}
  let n=0;
  msgs.forEach(m=>{const hit=m.textContent.toLowerCase().includes(t);m.classList.toggle('hidden',!hit);if(hit)n++;});
  days.forEach(d=>{let has=false,e=d.nextElementSibling;
    while(e&&!e.classList.contains('day')){if(e.classList.contains('msg')&&!e.classList.contains('hidden')){has=true;break;}e=e.nextElementSibling;}
    d.classList.toggle('hidden',!has);});
  c.textContent=n+' match'+(n===1?'':'es');
});
document.addEventListener('click',e=>{if(e.target.classList.contains('shot'))e.target.classList.toggle('zoom');});
</script>
</body></html>
'@)

    $htmlPath = Join-Path $runDir 'transcript.html'
    [System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
    $mb = [math]::Round((Get-Item $htmlPath).Length / 1MB, 1)
    Write-Host "Wrote transcript.html ($mb MB)" -ForegroundColor Green
    if ($mb -gt 250 -and -not $HtmlLinkImages) {
        Write-Warning "That's a big single file. Re-run with -HtmlLinkImages for a small HTML that references the images folder instead."
    }
    Write-Host "For a PDF: open it and Ctrl+P, or run" -ForegroundColor DarkGray
    Write-Host "  msedge --headless --disable-gpu --print-to-pdf=`"$runDir\transcript.pdf`" `"$htmlPath`"" -ForegroundColor DarkGray
}

# ---- manifest ----
$roster = $script:SenderRoster.GetEnumerator() | Sort-Object { -$_.Value.Count }
$rosterLines = foreach ($e in $roster) {
    "| {0} | {1} | {2} |" -f $e.Value.Name, $(if ($e.Value.Guid) { $e.Value.Guid } else { 'n/a' }), $e.Value.Count
}
$chunkLines = foreach ($f in $chunkFiles) { "- ``chunks/$f``" }
$imgLines   = foreach ($x in $imgIndex) { "| ``{0}`` | {1} | {2} | {3} KB |" -f $x.File, $x.When, $x.Sender, [math]::Round($x.Bytes/1KB) }
$totalChars = ($records | ForEach-Object { $_.text.Length } | Measure-Object -Sum).Sum

@"
# Export manifest

- **Chat id:** ``$ChatId``
- **Participants:** $participants
- **Exported by:** $($ctx.Account) on $(Get-Date -Format 'yyyy-MM-dd HH:mm')
- **Date range:** $($first.ToString('yyyy-MM-dd')) to $($last.ToString('yyyy-MM-dd'))
- **Messages retrieved / kept:** $($raw.Count) / $($records.Count)
- **Images downloaded:** $imgCounter
- **Approx. text volume:** $totalChars chars (~$([math]::Round($totalChars/4000)) k tokens)

## Senders found

| Name | Sender GUID | Messages |
|------|-------------|----------|
$($rosterLines -join "`n")

Anyone still showing as **Unknown User** can be named on a re-run:

``````powershell
.\Export-TeamsChat.ps1 -ChatId "$ChatId" -NameMap @{ '<guid above>' = 'Real Name' }
``````

## Filtered out

| Reason | Count |
|--------|-------|
| System events (joins, renames, calls) | $($skipped.System) |
| Deleted messages | $($skipped.Deleted) |
| Empty after cleanup | $($skipped.Empty) |

Recover with ``-KeepSystemMessages``, ``-KeepReactions``, ``-KeepFullUrls``.

## Image index

$(if ($imgIndex.Count) { "| File | Sent | Sender | Size |`n|------|------|--------|------|`n" + ($imgLines -join "`n") } else { "_No inline images, or -NoImages was used._" })

## Files

$($chunkLines -join "`n")
- ``transcript.jsonl`` - one JSON object per message
- ``images\`` - $imgCounter pasted screenshots
- ``images\index.md`` - what each screenshot is, who sent it, and the conversation either side of it
"@ | Out-File (Join-Path $runDir 'manifest.md') -Encoding utf8

Write-Host "Wrote manifest.md" -ForegroundColor Green
Write-Host "`nOutput: $runDir" -ForegroundColor Cyan

if ($script:Interactive -and $script:EquivalentCommand) {
    Write-Host "`nTo repeat this without the prompts:" -ForegroundColor DarkGray
    Write-Host "  $($script:EquivalentCommand)" -ForegroundColor DarkGray
}

$unknown = $roster | Where-Object { $_.Value.Name -eq 'Unknown User' }
if ($unknown) {
    Write-Host "`nUnresolved sender(s) - re-run with -NameMap to fix:" -ForegroundColor Yellow
    foreach ($e in $unknown) {
        Write-Host ("  -NameMap @{{ '{0}' = 'Their Name' }}   ({1} messages)" -f $e.Value.Guid, $e.Value.Count) -ForegroundColor Yellow
    }
}
Write-Host ""