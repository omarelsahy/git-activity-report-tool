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
    [switch] $SkipPdf
)

$ErrorActionPreference = "Stop"

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
        [string[]] $Args
    )
    Push-Location $RepoPath
    try {
        return @(& git @Args 2>$null)
    }
    finally {
        Pop-Location
    }
}

function Invoke-GitCount {
    param(
        [string] $RepoPath,
        [string[]] $Args
    )
    $lines = Invoke-GitLines -RepoPath $RepoPath -Args $Args
    if (-not $lines -or -not $lines[0]) { return 0 }
    $count = 0
    [void][int]::TryParse($lines[0].Trim(), [ref]$count)
    return $count
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

    $countBase = @("--since=$SinceText", "--until=$UntilText") + $authorArgs + @("HEAD")
    $total = Invoke-GitCount -RepoPath $RepoPath -Args (@("rev-list", "--count") + $countBase)
    $merges = Invoke-GitCount -RepoPath $RepoPath -Args (@("rev-list", "--count", "--merges") + $countBase)
    $nonMerges = Invoke-GitCount -RepoPath $RepoPath -Args (@("rev-list", "--count", "--no-merges") + $countBase)

    $numstatArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=tformat:", "--numstat")
    $numstatLines = Invoke-GitLines -RepoPath $RepoPath -Args $numstatArgs
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
    foreach ($line in (Invoke-GitLines -RepoPath $RepoPath -Args $nameOnlyArgs)) {
        if ($line -and $line.Trim()) {
            [void]$files.Add($line.Trim())
        }
    }

    $subjectArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=%s")
    $hashArgs = @("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--no-merges", "--pretty=%H")
    $subjects = Invoke-GitLines -RepoPath $RepoPath -Args $subjectArgs
    $hashes = Invoke-GitLines -RepoPath $RepoPath -Args $hashArgs

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

    $authors = Invoke-GitLines -RepoPath $RepoPath -Args (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--format=%aN")) | Sort-Object -Unique
    $highlights = Invoke-GitLines -RepoPath $RepoPath -Args (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--oneline", "-n", "$MaxHighlights"))
    $highlightTotal = (Invoke-GitLines -RepoPath $RepoPath -Args (@("log", "--since=$SinceText", "--until=$UntilText") + $authorArgs + @("--oneline"))).Count

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
        return @((Resolve-Path -LiteralPath $top[0]).Path)
    }

    $children = Get-ChildItem -LiteralPath $resolved -Directory
    $repos = @()
    foreach ($child in $children) {
        if ($child.Name -eq "_sync-logs") { continue }
        if (Test-Path -LiteralPath (Join-Path $child.FullName ".git")) {
            $repos += $child.FullName
        }
    }
    return $repos
}

$bounds = Get-PeriodBounds -Mode $Period -SinceInput $Since -UntilInput $Until
$sinceDate = [datetime]$bounds.Since
$untilDate = [datetime]$bounds.Until
$sinceText = $sinceDate.ToString("yyyy-MM-dd HH:mm:ss")
$untilText = $untilDate.ToString("yyyy-MM-dd HH:mm:ss")
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

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$defaultName = if ($ReportName) {
    $ReportName
}
else {
    "Git-$label-Report-" + $sinceDate.ToString("MM-dd-yyyy")
}

$pdfPath = Join-Path $OutputDirectory ($defaultName + ".pdf")
if (Test-Path -LiteralPath $pdfPath) {
    $pdfPath = Join-Path $OutputDirectory ($defaultName + "-" + (Get-Date -Format "HHmmss") + ".pdf")
}
$mdPath = [System.IO.Path]::ChangeExtension($pdfPath, ".md")

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("# Git activity report")
[void]$sb.AppendLine()
[void]$sb.AppendLine("**Report generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') local  ")
[void]$sb.AppendLine("**Period type:** $label  ")
[void]$sb.AppendLine("**Period:** **$sinceText** through **$untilText**")
if ($Author) {
    [void]$sb.AppendLine("**Author filter:** --author=""$Author""")
}
[void]$sb.AppendLine()
[void]$sb.AppendLine("**Scope:** $($repoTargets.Count) discovered repositories; $($sortedMetrics.Count) included in output.")
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
        [void]$sb.AppendLine()
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
        HeaderLabel = $HeaderLabel
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

