# Git Activity Report Tool

Generate daily or weekly Git activity reports across one repository or a folder of sibling repositories, then export to PDF.

## Features

- Scans a single repo or a container folder with multiple repos.
- Computes commit counts, churn (`+/-`), files touched, conventional commit signals, authors, and highlights.
- Excludes repos with no commits by default.
- Sorts per-repo sections by change volume (`added + removed`), then commit count.
- Exports a styled PDF using Edge headless printing (Pandoc-enhanced HTML when available).
- Optional `--author` filter to report on one contributor only.
- Optional markdown-only mode (`-SkipPdf`) for CI environments.
- Trello task section support:
  - always: detects Trello-like task refs from commit messages
  - optional: loads live Trello board data when API credentials are provided
  - per board: **To Do created** and **entered In Progress** and **completed** counts for the report window, plus **current list totals** in parentheses (styled in PDF as period vs total)
  - Trello metrics are included per repository and aggregated in rollup totals

## Requirements

- Windows PowerShell 5.1+
- Git on `PATH`
- Microsoft Edge installed
- Optional: Pandoc on `PATH` (for higher quality markdown rendering)
- Optional (for live Trello data): Trello API key/token and board ID or board name

## Quick Start

```powershell
Set-Location "C:\path\to\git-activity-report-tool"

.\scripts\new-git-activity-report.ps1 `
  -RootPath "C:\path\to\repos" `
  -Period daily `
  -OutputDirectory "C:\path\to\reports\daily"
```

## Author-Filtered Example

```powershell
.\scripts\new-git-activity-report.ps1 `
  -RootPath "C:\path\to\repos" `
  -Period weekly `
  -Author "Jane Doe" `
  -OutputDirectory "C:\path\to\reports\weekly"
```

## Custom Time Window Example

```powershell
.\scripts\new-git-activity-report.ps1 `
  -RootPath "C:\path\to\repos" `
  -Period custom `
  -Since "2026-04-01 00:00:00" `
  -Until "2026-04-08 00:00:00" `
  -OutputDirectory "C:\path\to\reports\custom"
```

## CI / Markdown-Only Example

```powershell
.\scripts\new-git-activity-report.ps1 `
  -RootPath "." `
  -Period custom `
  -Since "2026-04-01 00:00:00" `
  -Until "2026-04-02 00:00:00" `
  -OutputDirectory ".\Reports\CI" `
  -KeepMarkdown `
  -SkipPdf
```

## Trello Integration Example

```powershell
.\scripts\new-git-activity-report.ps1 `
  -RootPath "C:\path\to\repos" `
  -Period weekly `
  -OutputDirectory "C:\path\to\reports\weekly" `
  -TrelloApiKey "<key>" `
  -TrelloApiToken "<token>" `
  -TrelloBoardName "Engineering Board"
```

You can also set environment variables:

- `TRELLO_API_KEY`
- `TRELLO_API_TOKEN` (or `TRELLO_TOKEN` if the API token variable is unset)
- `TRELLO_BOARD_ID` (or `TRELLO_BOARD_NAME`)

If you use the local Trello MCP server under `%USERPROFILE%\tools\trello-mcp-server\`, the report script will **automatically load** `%USERPROFILE%\tools\trello-mcp-server\.env` when present (same keys as the MCP). Override with `-TrelloEnvFile` or `TRELLO_ENV_FILE`.

Board selection matches a repository folder name to a Trello board **by slug** (letters and digits only), so `sdx-lockbox-sensor` matches boards like **SDX Lockbox Sensor** or `sdx_lockbox_sensor`. List names like **To-Do** / **In-Progress** are normalized before matching workflow columns.

If no list titles match those patterns, the tool uses **column order** (`pos`): **leftmost** list is treated as intake (To Do), **rightmost** as done (Complete), and **lists in between** as In Progress for move metrics.

On Windows PowerShell 5.1, large Trello JSON arrays are sometimes deserialized as a **single object with parallel properties** (for example many `date` values in one property). The report script expands those shapes for **lists**, **cards**, and **board actions** so counts stay accurate.

## Output

The generator writes:

- Markdown report (intermediate)
- PDF report (final)

By default, the markdown file is deleted when PDF generation succeeds.

## Scripts

- `scripts\new-git-activity-report.ps1` - Builds report metrics and markdown, then invokes PDF conversion (use `-SkipPdf` for markdown-only).
- `scripts\convert-markdown-to-pdf.ps1` - Converts markdown to a styled PDF.

