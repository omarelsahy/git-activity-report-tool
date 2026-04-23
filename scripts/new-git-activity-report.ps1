#Requires -Version 5.1
<#
.SYNOPSIS
  Generate a git activity report (daily, weekly, or custom) and export to PDF.
#>
param(
    [Parameter(Mandatory = $false)]
    [string] $RootPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [ValidateSet("daily", "weekly", "custom")]
    [string] $Period = "daily",

    [Parameter(Mandatory = $false)]
    [datetime] $Since,

    [Parameter(Mandatory = $false)]
    [datetime] $Until,

    [Parameter(Mandatory = $false)]
    [string] $Author,

    [Parameter(Mandatory = $false)]
    [string] $AuthorDisplayName,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [Parameter(Mandatory = $false)]
    [string] $ReportName,

    [Parameter(Mandatory = $false)]
    [string] $HeaderLabel = "Git Activity Report",

    [Parameter(Mandatory = $false)]
    [string] $LogoPath,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeEmptyRepos,

    [Parameter(Mandatory = $false)]
    [switch] $KeepMarkdown,

    [Parameter(Mandatory = $false)]
    [switch] $SkipPdf,

    [Parameter(Mandatory = $false)]
    [string] $TrelloApiKey = $env:TRELLO_API_KEY,

    [Parameter(Mandatory = $false)]
    [string] $TrelloApiToken = $env:TRELLO_API_TOKEN,

    [Parameter(Mandatory = $false)]
    [string] $TrelloBoardId = $env:TRELLO_BOARD_ID,

    [Parameter(Mandatory = $false)]
    [string] $TrelloBoardName = $env:TRELLO_BOARD_NAME,

    [Parameter(Mandatory = $false)]
    [string] $TrelloEnvFile
)

$ErrorActionPreference = "Stop"

function Import-TrelloDotEnvFile {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $allowed = New-Object "System.Collections.Generic.HashSet[string]"
    foreach ($k in @("TRELLO_API_KEY", "TRELLO_TOKEN", "TRELLO_API_TOKEN", "TRELLO_BOARD_ID", "TRELLO_BOARD_NAME")) {
        [void]$allowed.Add($k)
    }
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith("#")) { continue }
        $eq = $t.IndexOf("=")
        if ($eq -lt 1) { continue }
        $k = $t.Substring(0, $eq).Trim()
        if (-not $allowed.Contains($k)) { continue }
        $v = $t.Substring($eq + 1).Trim()
        if (($v.Length -ge 2) -and (
                ($v.StartsWith('"') -and $v.EndsWith('"')) -or
                ($v.StartsWith("'") -and $v.EndsWith("'")))) {
            $v = $v.Substring(1, $v.Length - 2)
        }
        Set-Item -Path "Env:$k" -Value $v
    }
}

# Same credentials file the local Trello MCP server typically loads (cwd + .env).
# Do not call Join-Path with a null profile (Linux CI); USERPROFILE is Windows-specific.
$resolvedTrelloEnv = $null
$trelloEnvCandidates = @()
foreach ($x in @($TrelloEnvFile, $env:TRELLO_ENV_FILE)) {
    if (-not [string]::IsNullOrWhiteSpace($x)) {
        $trelloEnvCandidates += $x
    }
}
if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    $trelloEnvCandidates += (Join-Path $env:USERPROFILE "tools\trello-mcp-server\.env")
}

foreach ($candidate in $trelloEnvCandidates) {
    if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
    try {
        $full = [System.IO.Path]::GetFullPath($candidate)
    }
    catch {
        continue
    }
    if (-not [string]::IsNullOrWhiteSpace($full) -and (Test-Path -LiteralPath $full)) {
        $resolvedTrelloEnv = $full
        Import-TrelloDotEnvFile -Path $full
        break
    }
}

if (-not $TrelloApiKey) { $TrelloApiKey = $env:TRELLO_API_KEY }
if (-not $TrelloApiToken) { $TrelloApiToken = $env:TRELLO_API_TOKEN }
if (-not $TrelloApiToken) { $TrelloApiToken = $env:TRELLO_TOKEN }
if (-not $TrelloBoardId) { $TrelloBoardId = $env:TRELLO_BOARD_ID }
if (-not $TrelloBoardName) { $TrelloBoardName = $env:TRELLO_BOARD_NAME }

function Get-PeriodBounds {
    param(
        [string] $Mode,
        [Nullable[datetime]] $SinceInput,
        [Nullable[datetime]] $UntilInput
    )
    $now = Get-Date
    switch ($Mode) {
        "daily" {
            $start = Get-Date -Date $now.Date
            return @{ Since = $start; Until = $now; Label = "Daily"; MaxHighlights = 15 }
        }
        "weekly" {
            $dayOffset = (([int]$now.DayOfWeek + 6) % 7)
            $monday = (Get-Date -Date $now.Date).AddDays(-$dayOffset)
            return @{ Since = $monday; Until = $now; Label = "Weekly"; MaxHighlights = 25 }
        }
        "custom" {
            if (-not $SinceInput) { throw "When -Period custom is used, -Since is required." }
            $end = if ($UntilInput) { $UntilInput } else { $now }
            return @{ Since = $SinceInput; Until = $end; Label = "Custom"; MaxHighlights = 25 }
        }
        default { throw "Unsupported period: $Mode" }
    }
}

function Invoke-GitLines {
    param(
        [string] $RepoPath,
        [string[]] $GitArgs
    )
    Push-Location $RepoPath
    try {
        return @(& git @GitArgs 2>$null)
    }
    finally {
        Pop-Location
    }
}

function Invoke-GitCount {
    param(
        [string] $RepoPath,
        [string[]] $GitArgs
    )
    $lines = Invoke-GitLines -RepoPath $RepoPath -GitArgs $GitArgs
    if (-not $lines -or -not $lines[0]) { return 0 }
    $count = 0
    [void][int]::TryParse($lines[0].Trim(), [ref]$count)
    return $count
}

function Resolve-GitHubHttpsRepoUrl {
    param([string] $RemoteUrl)
    if ([string]::IsNullOrWhiteSpace($RemoteUrl)) { return $null }
    $u = $RemoteUrl.Trim()
    $built = $null
    if ($u -match '^git@github\.com:([^/]+)/([^\s/]+?)(?:\.git)?$') {
        $built = "https://github.com/$($matches[1])/$($matches[2])"
    }
    elseif ($u -match '^ssh://git@github\.com/([^/]+)/([^\s/]+?)(?:\.git)?$') {
        $built = "https://github.com/$($matches[1])/$($matches[2])"
    }
    elseif ($u -match '^https://github\.com/([^/]+)/([^\s/?#]+)') {
        $repoSeg = $matches[2] -replace '\.git$', ''
        $built = "https://github.com/$($matches[1])/$repoSeg"
    }
    if (-not $built) { return $null }
    return ($built -replace '\.git$', '')
}

function Get-GitHubRepoWebUrl {
    param([string] $RepoPath)
    $lines = @(Invoke-GitLines -RepoPath $RepoPath -GitArgs @("remote", "get-url", "origin"))
    if ($lines.Count -eq 0 -or [string]::IsNullOrWhiteSpace($lines[0])) { return $null }
    return (Resolve-GitHubHttpsRepoUrl -RemoteUrl $lines[0])
}

function Get-RepoMetrics {
    param(
        [string] $RepoPath,
        [string] $SinceText,
        [string] $UntilText,
        [string] $AuthorFilter,
        [int] $MaxHighlights
    )

    $authorArgs = @()
    if ($AuthorFilter) {
        $authorArgs = @("--author=$AuthorFilter")
    }

    # Counting hashes from git log is more resilient than rev-list --count when
    # git emits non-numeric output in some environments.
    $total = (Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--format=%H"))).Count
    $merges = (Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--merges", "--format=%H"))).Count
    $nonMerges = (Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--format=%H"))).Count

    $numstatArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=tformat:", "--numstat")
    $numstatLines = Invoke-GitLines -RepoPath $RepoPath -GitArgs $numstatArgs
    $added = 0L
    $removed = 0L
    $binaryRows = 0
    foreach ($line in $numstatLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split "`t", 3
        if ($parts.Count -lt 3) { continue }
        if ($parts[0] -eq "-" -or $parts[1] -eq "-") {
            $binaryRows++
            continue
        }
        $a = 0
        $d = 0
        [void][int]::TryParse($parts[0], [ref]$a)
        [void][int]::TryParse($parts[1], [ref]$d)
        $added += $a
        $removed += $d
    }

    $files = New-Object "System.Collections.Generic.HashSet[string]"
    $nameOnlyArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--name-only", "--pretty=format:")
    foreach ($line in (Invoke-GitLines -RepoPath $RepoPath -GitArgs $nameOnlyArgs)) {
        if ($line -and $line.Trim()) {
            [void]$files.Add($line.Trim())
        }
    }

    $subjectArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=%s")
    $hashArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=%H")
    $subjects = Invoke-GitLines -RepoPath $RepoPath -GitArgs $subjectArgs
    $hashes = Invoke-GitLines -RepoPath $RepoPath -GitArgs $hashArgs

    $feat = 0; $fix = 0; $perf = 0; $refactor = 0; $docs = 0; $test = 0; $chore = 0; $build = 0; $ci = 0
    $featHeur = New-Object "System.Collections.Generic.HashSet[string]"
    $fixHeur = New-Object "System.Collections.Generic.HashSet[string]"

    for ($i = 0; $i -lt $subjects.Count; $i++) {
        $subject = $subjects[$i]
        $hash = if ($i -lt $hashes.Count) { $hashes[$i] } else { "" }

        if ($subject -match "^(?i)feat(\([^)]*\))?:") { $feat++; continue }
        if ($subject -match "^(?i)fix(\([^)]*\))?:") { $fix++; continue }
        if ($subject -match "^(?i)perf(\([^)]*\))?:") { $perf++; continue }
        if ($subject -match "^(?i)refactor(\([^)]*\))?:") { $refactor++; continue }
        if ($subject -match "^(?i)docs(\([^)]*\))?:") { $docs++; continue }
        if ($subject -match "^(?i)test(\([^)]*\))?:") { $test++; continue }
        if ($subject -match "^(?i)chore(\([^)]*\))?:") { $chore++; continue }
        if ($subject -match "^(?i)build(\([^)]*\))?:") { $build++; continue }
        if ($subject -match "^(?i)ci(\([^)]*\))?:") { $ci++; continue }

        if ($subject -match "(?i)\b(add|implement|introduce)\b") { [void]$featHeur.Add($hash) }
        if ($subject -match "(?i)\b(fix|bug|regression|patch)\b") { [void]$fixHeur.Add($hash) }
    }

    $authors = Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--format=%aN")) | Sort-Object -Unique
    $highlights = Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--oneline", "-n", "$MaxHighlights"))
    $highlightTotal = (Invoke-GitLines -RepoPath $RepoPath -GitArgs (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--oneline"))).Count

    return [pscustomobject]@{
        Name = (Split-Path $RepoPath -Leaf)
        Path = $RepoPath
        Total = $total
        Merges = $merges
        NonMerges = $nonMerges
        Added = $added
        Removed = $removed
        BinaryRows = $binaryRows
        FileCount = $files.Count
        Feat = $feat
        Fix = $fix
        FeatHeur = $featHeur.Count
        FixHeur = $fixHeur.Count
        Perf = $perf
        Refactor = $refactor
        Docs = $docs
        Test = $test
        Chore = $chore
        Build = $build
        CI = $ci
        Authors = $authors
        Subjects = $subjects
        Highlights = $highlights
        HighlightTotal = $highlightTotal
        ChangeVolume = ($added + $removed)
    }
}

function Get-RepoTargets {
    param([string] $Path)

    $resolved = (Resolve-Path -LiteralPath $Path).Path

    Push-Location $resolved
    try {
        try {
            $top = @(& git rev-parse --show-toplevel 2>$null)
        }
        catch {
            $top = @()
        }
    }
    finally {
        Pop-Location
    }

    if ($top -and $top[0]) {
        # Leading comma: PowerShell unwraps single-element arrays from return values to a scalar;
        # without it, foreach over $repoTargets would iterate characters of the path string.
        return ,@((Resolve-Path -LiteralPath $top[0]).Path)
    }

    $children = Get-ChildItem -LiteralPath $resolved -Directory
    $repos = @()
    foreach ($child in $children) {
        if ($child.Name -eq "_sync-logs") { continue }
        if (Test-Path -LiteralPath (Join-Path $child.FullName ".git")) {
            $repos += $child.FullName
        }
    }
    return ,$repos
}

function Get-TrelloBoardByName {
    param(
        [string] $ApiKey,
        [string] $ApiToken,
        [string] $BoardName
    )
    $uri = "https://api.trello.com/1/members/me/boards?fields=id,name,url&key=$ApiKey&token=$ApiToken"
    $boards = Invoke-RestMethod -Method Get -Uri $uri
    if (-not $boards) { return $null }
    return @($boards | Where-Object { $_.name -eq $BoardName } | Select-Object -First 1)[0]
}

function Get-TrelloBoards {
    param(
        [string] $ApiKey,
        [string] $ApiToken
    )
    $uri = "https://api.trello.com/1/members/me/boards?fields=id,name,url,closed&key=$ApiKey&token=$ApiToken"
    return @(Invoke-RestMethod -Method Get -Uri $uri)
}

function Expand-TrelloListRecords {
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload -is [System.Array]) {
        if ($Payload.Count -eq 1 -and $Payload[0].name -is [System.Array]) {
            return @(Expand-TrelloListRecords -Payload $Payload[0])
        }
        return @($Payload)
    }
    $names = $Payload.name
    if ($names -is [System.Array]) {
        $ids = $Payload.id
        $closedVals = $Payload.closed
        $poses = $Payload.pos
        $out = @()
        for ($i = 0; $i -lt $names.Count; $i++) {
            $out += [pscustomobject]@{
                id = $ids[$i]
                name = [string]$names[$i]
                closed = $closedVals[$i]
                pos = $poses[$i]
            }
        }
        return $out
    }
    return @($Payload)
}

function Expand-TrelloCardRecords {
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload -is [System.Array]) {
        if ($Payload.Count -eq 1 -and $Payload[0].name -is [System.Array]) {
            return @(Expand-TrelloCardRecords -Payload $Payload[0])
        }
        return @($Payload)
    }
    $names = $Payload.name
    if ($names -is [System.Array]) {
        $ids = $Payload.id
        $idLists = $Payload.idList
        $closedVals = $Payload.closed
        $shortUrls = $Payload.shortUrl
        $out = @()
        for ($i = 0; $i -lt $names.Count; $i++) {
            $row = [ordered]@{
                id = $ids[$i]
                name = [string]$names[$i]
                idList = $idLists[$i]
                closed = $closedVals[$i]
            }
            if ($null -ne $shortUrls) {
                $arr = $shortUrls
                if ($arr -is [System.Array] -and $i -lt $arr.Count) {
                    $row["shortUrl"] = $arr[$i]
                }
            }
            $out += [pscustomobject]$row
        }
        return $out
    }
    return @($Payload)
}

function Expand-TrelloActionRecords {
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload -is [System.Array]) {
        if ($Payload.Count -eq 1 -and $Payload[0].date -is [System.Array]) {
            return @(Expand-TrelloActionRecords -Payload $Payload[0])
        }
        return @($Payload)
    }
    $dates = $Payload.date
    if ($dates -is [System.Array]) {
        $ids = $Payload.id
        $types = $Payload.type
        $dataArr = $Payload.data
        $out = @()
        for ($i = 0; $i -lt $dates.Count; $i++) {
            $out += [pscustomobject]@{
                id = $ids[$i]
                type = [string]$types[$i]
                date = [string]$dates[$i]
                data = $dataArr[$i]
            }
        }
        return $out
    }
    return @($Payload)
}

function Get-TrelloNameSlug {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return ($Name.ToLowerInvariant() -replace "[^a-z0-9]+", "")
}

function Normalize-TrelloListLabel {
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.ToLowerInvariant() -replace "[_\-]+", " "
    $t = $t -replace "\s+", " "
    return $t.Trim()
}

function Test-TrelloEntityOpen {
    param($Item)
    if ($null -eq $Item) { return $false }
    if ($null -eq $Item.closed) { return $true }
    if ($Item.closed -is [bool]) { return -not [bool]$Item.closed }
    $s = [string]$Item.closed
    return ($s -notmatch "^(?i)true$")
}

function Get-TrelloListPos {
    param($List)
    if ($null -eq $List -or $null -eq $List.pos) { return 0.0 }
    $p = $List.pos
    if ($p -is [System.Array]) {
        $p = $p | Select-Object -First 1
    }
    $d = 0.0
    [void][double]::TryParse([string]$p, [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$d)
    return $d
}

function Format-TrelloPeriodTotalHtml {
    param(
        [int] $PeriodCount,
        [int] $BoardTotal
    )
    return ('<span class="trello-recent">{0}</span><span class="trello-total"> ({1})</span>' -f $PeriodCount, $BoardTotal)
}

function Parse-TrelloTimestamp {
    param([string] $Iso)
    if ([string]::IsNullOrWhiteSpace($Iso)) { return $null }
    try {
        return [datetimeoffset]::Parse($Iso, $null, [System.Globalization.DateTimeStyles]::RoundtripKind).UtcDateTime
    }
    catch {
        return $null
    }
}

function Resolve-TrelloWorkflowLists {
    param(
        [object[]] $Lists
    )
    $openLists = @($Lists | Where-Object { Test-TrelloEntityOpen -Item $_ })
    $byNorm = @{}
    foreach ($l in $openLists) {
        $key = Normalize-TrelloListLabel -Text $l.name
        if (-not $key) { continue }
        if (-not $byNorm.ContainsKey($key)) { $byNorm[$key] = @() }
        $byNorm[$key] += $l
    }

    $todo = $null
    foreach ($k in @("to do", "todo", "backlog", "new", "queue", "ideas")) {
        if ($byNorm.ContainsKey($k)) {
            $todo = @($byNorm[$k]) | Select-Object -First 1
            break
        }
    }
    if (-not $todo) {
        $todo = $openLists | Where-Object { (Normalize-TrelloListLabel -Text $_.name) -match "^(to do|todo|backlog)$" } | Select-Object -First 1
    }

    $inProgress = $null
    foreach ($k in @("in progress", "doing", "active", "wip")) {
        if ($byNorm.ContainsKey($k)) {
            $inProgress = @($byNorm[$k]) | Select-Object -First 1
            break
        }
    }
    if (-not $inProgress) {
        $inProgress = $openLists |
            Where-Object {
                $nl = Normalize-TrelloListLabel -Text $_.name
                $nl -eq "in progress" -or $nl -eq "doing" -or $nl -eq "active" -or $nl -eq "wip"
            } |
            Select-Object -First 1
    }

    $complete = $null
    foreach ($k in @("complete", "completed", "done")) {
        if ($byNorm.ContainsKey($k)) {
            $complete = @($byNorm[$k]) | Select-Object -First 1
            break
        }
    }
    if (-not $complete) {
        $complete = $openLists | Where-Object { (Normalize-TrelloListLabel -Text $_.name) -match "^(complete|completed|done)$" } | Select-Object -First 1
    }

    return [pscustomobject]@{
        TodoList = $todo
        InProgressList = $inProgress
        CompleteList = $complete
    }
}

function Get-TrelloBoardActionsPaged {
    param(
        [string] $ApiKey,
        [string] $ApiToken,
        [string] $BoardId,
        [datetime] $SinceUtc,
        [datetime] $UntilUtc
    )

    $sinceIso = [System.Uri]::EscapeDataString($SinceUtc.ToString("o"))
    $collected = New-Object "System.Collections.Generic.List[object]"
    $beforeId = $null
    $safety = 0

    while ($safety -lt 200) {
        $safety++
        $uri = "https://api.trello.com/1/boards/$BoardId/actions?limit=1000&filter=createCard,copyCard,updateCard&since=$sinceIso&fields=id,type,date,data&key=$ApiKey&token=$ApiToken"
        if ($beforeId) {
            $uri += "&before=$beforeId"
        }

        $batch = Expand-TrelloActionRecords -Payload (Invoke-RestMethod -Method Get -Uri $uri)
        if (-not $batch -or $batch.Count -eq 0) { break }

        $oldestUtc = $null
        foreach ($a in $batch) {
            $ad = Parse-TrelloTimestamp -Iso $a.date
            if (-not $ad) { continue }
            if ($ad -gt $UntilUtc) { continue }
            if ($ad -ge $SinceUtc -and $ad -le $UntilUtc) {
                [void]$collected.Add($a)
            }
            if (-not $oldestUtc -or $ad -lt $oldestUtc) { $oldestUtc = $ad }
        }

        if ($batch.Count -lt 1000) { break }
        if ($oldestUtc -and $oldestUtc -lt $SinceUtc) { break }

        $beforeId = $batch[-1].id
        if (-not $beforeId) { break }
    }

    return ,($collected.ToArray())
}

function Get-TrelloCardUrl {
    param([object] $CardData)
    if (-not $CardData) { return $null }
    if ($CardData.shortUrl) { return [string]$CardData.shortUrl }
    if ($CardData.shortLink) { return "https://trello.com/c/$($CardData.shortLink)" }
    return $null
}

function Get-TrelloBoardWorkMetrics {
    param(
        [string] $ApiKey,
        [string] $ApiToken,
        [string] $BoardId,
        [datetime] $SinceDate,
        [datetime] $UntilDate
    )

    $sinceUtc = $SinceDate.ToUniversalTime()
    $untilUtc = $UntilDate.ToUniversalTime()

    $listsUri = "https://api.trello.com/1/boards/$BoardId/lists?fields=id,name,closed,pos&key=$ApiKey&token=$ApiToken"
    $cardsUri = "https://api.trello.com/1/boards/$BoardId/cards?fields=id,name,idList,shortUrl,closed&key=$ApiKey&token=$ApiToken"
    $lists = Expand-TrelloListRecords -Payload (Invoke-RestMethod -Method Get -Uri $listsUri)
    $cards = Expand-TrelloCardRecords -Payload (Invoke-RestMethod -Method Get -Uri $cardsUri)

    $workflow = Resolve-TrelloWorkflowLists -Lists $lists

    $openListsOrdered = @(
        $lists |
        Where-Object { Test-TrelloEntityOpen -Item $_ } |
        Sort-Object @{ Expression = { Get-TrelloListPos -List $_ }; Ascending = $true }
    )

    $todoList = $workflow.TodoList
    $completeList = $workflow.CompleteList
    $inProgressListNamed = $workflow.InProgressList

    if (-not $todoList -and $openListsOrdered.Count -ge 1) {
        $todoList = $openListsOrdered[0]
    }
    if (-not $completeList -and $openListsOrdered.Count -ge 2) {
        $completeList = $openListsOrdered[-1]
    }

    $inProgressTargetIds = @()
    if ($inProgressListNamed) {
        $inProgressTargetIds = @([string]$inProgressListNamed.id)
    }
    elseif ($todoList -and $completeList) {
        $orderedIds = @($openListsOrdered | ForEach-Object { [string]$_.id })
        $ti = [array]::IndexOf($orderedIds, [string]$todoList.id)
        $ci = [array]::IndexOf($orderedIds, [string]$completeList.id)
        if ($ti -ge 0 -and $ci -gt $ti) {
            foreach ($mid in $openListsOrdered[($ti + 1)..($ci - 1)]) {
                $inProgressTargetIds += [string]$mid.id
            }
        }
    }

    $todoId = if ($todoList) { [string]$todoList.id } else { $null }
    $completeId = if ($completeList) { [string]$completeList.id } else { $null }

    $openCards = @($cards | Where-Object { Test-TrelloEntityOpen -Item $_ })
    $todoTotal = if ($todoId) { @($openCards | Where-Object { [string]$_.idList -eq $todoId }).Count } else { 0 }
    $inProgressTotal = 0
    foreach ($ipId in $inProgressTargetIds) {
        $inProgressTotal += @($openCards | Where-Object { [string]$_.idList -eq $ipId }).Count
    }
    $completeTotal = if ($completeId) { @($openCards | Where-Object { [string]$_.idList -eq $completeId }).Count } else { 0 }

    $todoCreated = 0
    $enteredInProgress = 0
    $completedMoves = 0
    $activitySeen = @{}
    $activityRows = New-Object "System.Collections.Generic.List[object]"

    $actions = Get-TrelloBoardActionsPaged -ApiKey $ApiKey -ApiToken $ApiToken -BoardId $BoardId -SinceUtc $sinceUtc -UntilUtc $untilUtc

    foreach ($a in $actions) {
        $ad = Parse-TrelloTimestamp -Iso $a.date
        if (-not $ad) { continue }

        if ($a.type -eq "createCard" -or $a.type -eq "copyCard") {
            $listId = $null
            if ($a.data -and $a.data.list -and $a.data.list.id) { $listId = [string]$a.data.list.id }
            if ($todoId -and $listId -eq $todoId) {
                $todoCreated++
            }

            $card = $a.data.card
            if ($card) {
                $kindLabel = if ($a.type -eq "copyCard") { "Copied" } else { "Created" }
                $key = "c:$([string]$card.id)"
                if (-not $activitySeen.ContainsKey($key)) {
                    $activitySeen[$key] = $true
                    $activityRows.Add([pscustomobject]@{
                            WhenUtc = $ad
                            Kind = $kindLabel
                            Name = [string]$card.name
                            Url = (Get-TrelloCardUrl -CardData $card)
                        })
                }
            }
            continue
        }

        if ($a.type -eq "updateCard") {
            $listAfter = $null
            $listBefore = $null
            if ($a.data) {
                if ($a.data.listAfter -and $a.data.listAfter.id) { $listAfter = [string]$a.data.listAfter.id }
                if ($a.data.listBefore -and $a.data.listBefore.id) { $listBefore = [string]$a.data.listBefore.id }
            }

            if ($inProgressTargetIds.Count -gt 0 -and $listAfter -and ($inProgressTargetIds -contains $listAfter) -and
                $listBefore -and ($inProgressTargetIds -notcontains $listBefore)) {
                $enteredInProgress++
            }

            if ($completeId -and $listAfter -eq $completeId -and $listBefore -ne $completeId) {
                $completedMoves++
            }

            if ($listAfter -and $listBefore -and $listAfter -ne $listBefore) {
                $card = $a.data.card
                if ($card) {
                    $key = "m:$([string]$card.id):$($ad.Ticks):$listAfter"
                    if (-not $activitySeen.ContainsKey($key)) {
                        $activitySeen[$key] = $true
                        $activityRows.Add([pscustomobject]@{
                                WhenUtc = $ad
                                Kind = "Moved"
                                Name = [string]$card.name
                                Url = (Get-TrelloCardUrl -CardData $card)
                            })
                    }
                }
            }
        }
    }

    $recentActivity = @($activityRows | Sort-Object -Property WhenUtc -Descending | Select-Object -First 25)

    $todoLabel = if ($todoList) { [string]$todoList.name } else { "—" }
    if ($inProgressListNamed) {
        $inProgLabel = [string]$inProgressListNamed.name
    }
    elseif ($inProgressTargetIds.Count -gt 0) {
        $midNames = @(
            $openListsOrdered |
            Where-Object { $inProgressTargetIds -contains [string]$_.id } |
            ForEach-Object { $_.name }
        )
        $inProgLabel = ($midNames -join " · ")
    }
    else {
        $inProgLabel = "—"
    }
    $completeLabel = if ($completeList) { [string]$completeList.name } else { "—" }
    $resolvedListsSummary = "To Do column: **$todoLabel** · In Progress: **$inProgLabel** · Complete: **$completeLabel**"

    return [pscustomobject]@{
        TodoListName = if ($todoList) { [string]$todoList.name } else { $null }
        InProgressListName = if ($inProgressListNamed) { [string]$inProgressListNamed.name } elseif ($inProgLabel -ne "—") { $inProgLabel } else { $null }
        CompleteListName = if ($completeList) { [string]$completeList.name } else { $null }
        ResolvedListsSummary = $resolvedListsSummary
        TodoCreatedInPeriod = $todoCreated
        TodoTotalOnBoard = $todoTotal
        InProgressEnteredInPeriod = $enteredInProgress
        InProgressTotalOnBoard = $inProgressTotal
        CompletedInPeriod = $completedMoves
        CompleteTotalOnBoard = $completeTotal
        RecentActivity = $recentActivity
        TodoHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $todoCreated -BoardTotal $todoTotal)
        InProgressHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $enteredInProgress -BoardTotal $inProgressTotal)
        CompleteHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $completedMoves -BoardTotal $completeTotal)
    }
}

$bounds = Get-PeriodBounds -Mode $Period -SinceInput $Since -UntilInput $Until
$sinceDate = [datetime]$bounds.Since
$untilDate = [datetime]$bounds.Until
$sinceText = $sinceDate.ToString("yyyy-MM-dd HH:mm:ss")
$untilText = $untilDate.ToString("yyyy-MM-dd HH:mm:ss")
$sinceDisplay = $sinceDate.ToString("MM/dd/yyyy")
$untilDisplay = $untilDate.ToString("MM/dd/yyyy")
$maxHighlights = [int]$bounds.MaxHighlights
$label = [string]$bounds.Label

$repoTargets = Get-RepoTargets -Path $RootPath
if (-not $repoTargets -or $repoTargets.Count -eq 0) {
    throw "No git repositories found under: $RootPath"
}

$metrics = @()
foreach ($repo in $repoTargets) {
    $metrics += Get-RepoMetrics -RepoPath $repo -SinceText $sinceText -UntilText $untilText -AuthorFilter $Author -MaxHighlights $maxHighlights
}

if (-not $IncludeEmptyRepos) {
    $metrics = @($metrics | Where-Object { $_.Total -gt 0 })
}

$sortedMetrics = @($metrics | Sort-Object @{ Expression = "ChangeVolume"; Descending = $true }, @{ Expression = "Total"; Descending = $true }, @{ Expression = "Name"; Descending = $false })

$rollupCommits = ($sortedMetrics | Measure-Object -Property Total -Sum).Sum
$rollupMerges = ($sortedMetrics | Measure-Object -Property Merges -Sum).Sum
$rollupNonMerges = ($sortedMetrics | Measure-Object -Property NonMerges -Sum).Sum
$rollupAdded = ($sortedMetrics | Measure-Object -Property Added -Sum).Sum
$rollupRemoved = ($sortedMetrics | Measure-Object -Property Removed -Sum).Sum
$rollupFiles = ($sortedMetrics | Measure-Object -Property FileCount -Sum).Sum
$rollupFeat = ($sortedMetrics | Measure-Object -Property Feat -Sum).Sum
$rollupFix = ($sortedMetrics | Measure-Object -Property Fix -Sum).Sum
$rollupFeatHeur = ($sortedMetrics | Measure-Object -Property FeatHeur -Sum).Sum
$rollupFixHeur = ($sortedMetrics | Measure-Object -Property FixHeur -Sum).Sum

$allAuthors = @($sortedMetrics | ForEach-Object { $_.Authors } | Sort-Object -Unique)

$trelloByRepo = @{}
$trelloRollupTodoPeriod = 0
$trelloRollupTodoTotal = 0
$trelloRollupInProgressPeriod = 0
$trelloRollupInProgressTotal = 0
$trelloRollupCompletePeriod = 0
$trelloRollupCompleteTotal = 0

if ($TrelloApiKey -and $TrelloApiToken -and $sortedMetrics.Count -gt 0) {
    try {
        $trelloBoards = Get-TrelloBoards -ApiKey $TrelloApiKey -ApiToken $TrelloApiToken
        $activeBoards = @($trelloBoards | Where-Object { Test-TrelloEntityOpen -Item $_ })
        $boardById = @{}
        foreach ($b in $activeBoards) { $boardById[$b.id] = $b }

        foreach ($repo in $sortedMetrics) {
            $selectedBoard = $null

            if ($TrelloBoardId -and $boardById.ContainsKey($TrelloBoardId)) {
                if ($sortedMetrics.Count -eq 1) {
                    $selectedBoard = $boardById[$TrelloBoardId]
                }
                else {
                    $bidBoard = $boardById[$TrelloBoardId]
                    if ($bidBoard.name -ieq $repo.Name -or
                        (Get-TrelloNameSlug -Name $bidBoard.name) -eq (Get-TrelloNameSlug -Name $repo.Name)) {
                        $selectedBoard = $bidBoard
                    }
                }
            }

            if (-not $selectedBoard -and $TrelloBoardName) {
                $nameSlug = Get-TrelloNameSlug -Name $TrelloBoardName
                $selectedBoard = @($activeBoards | Where-Object {
                        $_.name -ieq $TrelloBoardName -or (Get-TrelloNameSlug -Name $_.name) -eq $nameSlug
                    } | Select-Object -First 1)[0]
            }

            if (-not $selectedBoard) {
                $repoSlug = Get-TrelloNameSlug -Name $repo.Name
                $selectedBoard = @($activeBoards | Where-Object { (Get-TrelloNameSlug -Name $_.name) -eq $repoSlug } | Select-Object -First 1)[0]
            }

            if (-not $selectedBoard) {
                $selectedBoard = @($activeBoards | Where-Object { $_.name -ieq $repo.Name } | Select-Object -First 1)[0]
            }

            if ($selectedBoard) {
                try {
                    $summary = Get-TrelloBoardWorkMetrics -ApiKey $TrelloApiKey -ApiToken $TrelloApiToken -BoardId $selectedBoard.id -SinceDate $sinceDate -UntilDate $untilDate
                    $trelloByRepo[$repo.Name] = [pscustomobject]@{
                        BoardId = $selectedBoard.id
                        BoardName = $selectedBoard.name
                        BoardUrl = $selectedBoard.url
                        Summary = $summary
                    }

                    $trelloRollupTodoPeriod += [int]$summary.TodoCreatedInPeriod
                    $trelloRollupTodoTotal += [int]$summary.TodoTotalOnBoard
                    $trelloRollupInProgressPeriod += [int]$summary.InProgressEnteredInPeriod
                    $trelloRollupInProgressTotal += [int]$summary.InProgressTotalOnBoard
                    $trelloRollupCompletePeriod += [int]$summary.CompletedInPeriod
                    $trelloRollupCompleteTotal += [int]$summary.CompleteTotalOnBoard
                }
                catch {
                    Write-Warning "Failed to load Trello cards for repo '$($repo.Name)': $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to load Trello boards: $($_.Exception.Message)"
    }
}

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$defaultName = if ($ReportName) {
    $ReportName
}
else {
    if ($Period.ToLowerInvariant() -eq "weekly") {
        "Git-Weekly-Report-{0}_to_{1}" -f $sinceDate.ToString("MM-dd-yyyy"), $untilDate.ToString("MM-dd-yyyy")
    }
    else {
        "Git-$label-Report-" + $sinceDate.ToString("MM-dd-yyyy")
    }
}

$pdfPath = Join-Path $OutputDirectory ($defaultName + ".pdf")
if (Test-Path -LiteralPath $pdfPath) {
    $pdfPath = Join-Path $OutputDirectory ($defaultName + "-" + (Get-Date -Format "HHmmss") + ".pdf")
}
$mdPath = [System.IO.Path]::ChangeExtension($pdfPath, ".md")

$headerAuthor = $AuthorDisplayName
if ([string]::IsNullOrWhiteSpace($headerAuthor)) {
    if (-not [string]::IsNullOrWhiteSpace($Author)) {
        $headerAuthor = $Author
    }
    elseif ($repoTargets.Count -gt 0) {
        $gitNameLines = @(Invoke-GitLines -RepoPath $repoTargets[0] -GitArgs @("config", "user.name"))
        if ($gitNameLines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($gitNameLines[0])) {
            $headerAuthor = $gitNameLines[0].Trim()
        }
    }
}
if ([string]::IsNullOrWhiteSpace($headerAuthor)) {
    $headerAuthor = "Contributor"
}

$reportKindLine = switch ($Period.ToLowerInvariant()) {
    "daily" { "Daily Report" }
    "weekly" { "Weekly Report" }
    "custom" { "Custom Report" }
    default { "$label Report" }
}

$headerDateLine = if ($sinceDisplay -eq $untilDisplay) {
    $sinceDisplay
}
else {
    "$sinceDisplay - $untilDisplay"
}

$repoLinkParts = @()
foreach ($r in $sortedMetrics) {
    $web = Get-GitHubRepoWebUrl -RepoPath $r.Path
    $safeName = [string]$r.Name -replace '\]', '\]'
    if ($web) {
        $repoLinkParts += "[$safeName]($web)"
    }
    else {
        $repoLinkParts += $safeName
    }
}
$reposLine = if ($repoLinkParts.Count -eq 0) {
    "Repos: _(none included)_"
}
else {
    "Repos: $($repoLinkParts -join ', ')"
}

$pdfHeaderLabel = if ($PSBoundParameters.ContainsKey("HeaderLabel") -and -not [string]::IsNullOrWhiteSpace($HeaderLabel)) {
    $HeaderLabel
}
else {
    "$headerAuthor — $reportKindLine"
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine($headerAuthor)
[void]$sb.AppendLine()
[void]$sb.AppendLine($reportKindLine)
[void]$sb.AppendLine()
[void]$sb.AppendLine($headerDateLine)
[void]$sb.AppendLine()
[void]$sb.AppendLine($reposLine)
[void]$sb.AppendLine()
[void]$sb.AppendLine("---")
[void]$sb.AppendLine()
[void]$sb.AppendLine("## Rollup totals")
[void]$sb.AppendLine()
[void]$sb.AppendLine("| Metric | Total |")
[void]$sb.AppendLine("|--------|------:|")
[void]$sb.AppendLine("| Commits | $rollupCommits |")
[void]$sb.AppendLine("| Merge commits | $rollupMerges |")
[void]$sb.AppendLine("| Non-merge commits | $rollupNonMerges |")
[void]$sb.AppendLine("| Lines added (non-merge) | $rollupAdded |")
[void]$sb.AppendLine("| Lines removed (non-merge) | $rollupRemoved |")
[void]$sb.AppendLine("| Distinct files touched (sum by repo) | $rollupFiles |")
[void]$sb.AppendLine("| feat commits | $rollupFeat |")
[void]$sb.AppendLine("| fix commits | $rollupFix |")
[void]$sb.AppendLine("| Possible features (heuristic) | $rollupFeatHeur |")
[void]$sb.AppendLine("| Possible bug fixes (heuristic) | $rollupFixHeur |")
[void]$sb.AppendLine("| Trello boards matched to repos | $($trelloByRepo.Keys.Count) |")
$trelloRollupTodoHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $trelloRollupTodoPeriod -BoardTotal $trelloRollupTodoTotal)
$trelloRollupInProgressHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $trelloRollupInProgressPeriod -BoardTotal $trelloRollupInProgressTotal)
$trelloRollupCompleteHtml = (Format-TrelloPeriodTotalHtml -PeriodCount $trelloRollupCompletePeriod -BoardTotal $trelloRollupCompleteTotal)
[void]$sb.AppendLine("| Trello To Do created *(period / list total)* | $trelloRollupTodoHtml |")
[void]$sb.AppendLine("| Trello entered In Progress *(period / list total)* | $trelloRollupInProgressHtml |")
[void]$sb.AppendLine("| Trello completed *(moves to done list / list total)* | $trelloRollupCompleteHtml |")
[void]$sb.AppendLine()
[void]$sb.AppendLine("**Authors (unique):** $(if ($allAuthors.Count -gt 0) { $allAuthors -join '; ' } else { 'None' })")
[void]$sb.AppendLine()
[void]$sb.AppendLine("---")
[void]$sb.AppendLine()
[void]$sb.AppendLine("## Per-repository sections (sorted by change volume)")
[void]$sb.AppendLine()

if ($sortedMetrics.Count -eq 0) {
    [void]$sb.AppendLine("No commits found in the selected window.")
    [void]$sb.AppendLine()
}
else {
    foreach ($repo in $sortedMetrics) {
        $repoTrello = $null
        if ($trelloByRepo.ContainsKey($repo.Name)) {
            $repoTrello = $trelloByRepo[$repo.Name]
        }

        [void]$sb.AppendLine("### $($repo.Name)")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("| Metric | Value |")
        [void]$sb.AppendLine("|--------|------:|")
        [void]$sb.AppendLine("| Repository root | $($repo.Path) |")
        [void]$sb.AppendLine("| Commits / merges / non-merges | $($repo.Total) / $($repo.Merges) / $($repo.NonMerges) |")
        [void]$sb.AppendLine("| Lines added / removed (non-merge) | +$($repo.Added) / -$($repo.Removed) |")
        [void]$sb.AppendLine("| Change volume (added + removed) | $($repo.ChangeVolume) |")
        [void]$sb.AppendLine("| Distinct files touched | $($repo.FileCount) |")
        [void]$sb.AppendLine("| Binary --numstat rows ignored | $($repo.BinaryRows) |")
        [void]$sb.AppendLine("| feat / fix | $($repo.Feat) / $($repo.Fix) |")
        [void]$sb.AppendLine("| Possible features / bug fixes | $($repo.FeatHeur) / $($repo.FixHeur) |")
        [void]$sb.AppendLine("| perf / refactor / docs / test / chore / build / ci | $($repo.Perf) / $($repo.Refactor) / $($repo.Docs) / $($repo.Test) / $($repo.Chore) / $($repo.Build) / $($repo.CI) |")
        if ($repoTrello) {
            $s = $repoTrello.Summary
            [void]$sb.AppendLine("| Trello board | [$($repoTrello.BoardName)]($($repoTrello.BoardUrl)) |")
            [void]$sb.AppendLine("| Trello columns *(resolved)* | $($s.ResolvedListsSummary) |")
            [void]$sb.AppendLine("| Trello To Do created *(period / list total)* | $($s.TodoHtml) |")
            [void]$sb.AppendLine("| Trello entered In Progress *(period / list total)* | $($s.InProgressHtml) |")
            [void]$sb.AppendLine("| Trello completed *(moves to done list / list total)* | $($s.CompleteHtml) |")
        }
        [void]$sb.AppendLine()
        if ($repoTrello -and $repoTrello.Summary.RecentActivity.Count -gt 0) {
            [void]$sb.AppendLine("**Recent Trello activity (this period):**")
            foreach ($row in $repoTrello.Summary.RecentActivity) {
                if ($row.Url) {
                    $link = "[$($row.Name)]($($row.Url))"
                }
                else {
                    $link = [string]$row.Name
                }
                [void]$sb.AppendLine("- $($row.Kind): $link")
            }
            [void]$sb.AppendLine()
        }
        [void]$sb.AppendLine("**Highlights:**")
        if ($repo.Highlights.Count -gt 0) {
            foreach ($h in $repo.Highlights) {
                [void]$sb.AppendLine("- $h")
            }
            if ($repo.HighlightTotal -gt $repo.Highlights.Count) {
                $more = $repo.HighlightTotal - $repo.Highlights.Count
                [void]$sb.AppendLine()
                [void]$sb.AppendLine("*+ $more more commits in this window.*")
            }
        }
        else {
            [void]$sb.AppendLine("- None")
        }
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("---")
        [void]$sb.AppendLine()
    }
}

[void]$sb.AppendLine("## Caveats")
[void]$sb.AppendLine()
[void]$sb.AppendLine("- Metrics are derived from local git history only.")
[void]$sb.AppendLine("- Feature and bug counts are message-pattern estimates.")
[void]$sb.AppendLine("- Line churn is computed from non-merge commits using --numstat.")
[void]$sb.AppendLine("- Trello metrics are embedded into each repository section and rollup totals when Trello board mapping/credentials are available.")
[void]$sb.AppendLine("- Trello **period** counts use board actions in the report window (creates in the To Do list, moves into In Progress, moves into the completed list). **Totals in parentheses** are current open cards on that list.")
[void]$sb.AppendLine("- If list titles are not standard (To Do / In Progress / Done), the tool falls back to **board column order** (left = intake, right = done, columns between = in progress).")
[void]$sb.AppendLine()
[void]$sb.AppendLine("## Output")
[void]$sb.AppendLine()
[void]$sb.AppendLine("- PDF: $pdfPath")

Set-Content -LiteralPath $mdPath -Value $sb.ToString() -Encoding UTF8

if (-not $SkipPdf) {
    $converterPath = Join-Path $PSScriptRoot "convert-markdown-to-pdf.ps1"
    if (-not (Test-Path -LiteralPath $converterPath)) {
        throw "Missing converter script: $converterPath"
    }

    $convertArgs = @{
        MarkdownPath = $mdPath
        PdfPath = $pdfPath
        HeaderLabel = $pdfHeaderLabel
    }
    if ($LogoPath) {
        $convertArgs["LogoPath"] = $LogoPath
    }

    & $converterPath @convertArgs

    if (-not (Test-Path -LiteralPath $pdfPath)) {
        throw "PDF generation did not produce a file at: $pdfPath"
    }
}

if (-not $KeepMarkdown -and -not $SkipPdf) {
    Remove-Item -LiteralPath $mdPath -Force -ErrorAction SilentlyContinue
}

if ($SkipPdf) {
    Write-Host "Generated markdown report: $mdPath"
}
else {
    Write-Host "Generated PDF report: $pdfPath"
}

# Normalize process exit code for successful runs, even if earlier git probe commands failed non-fatally.
$global:LASTEXITCODE = 0


