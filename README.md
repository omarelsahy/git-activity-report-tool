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

## Requirements

- Windows PowerShell 5.1+
- Git on `PATH`
- Microsoft Edge installed
- Optional: Pandoc on `PATH` (for higher quality markdown rendering)

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

## Output

The generator writes:

- Markdown report (intermediate)
- PDF report (final)

By default, the markdown file is deleted when PDF generation succeeds.

## Scripts

- `scripts\new-git-activity-report.ps1` - Builds report metrics and markdown, then invokes PDF conversion.
- `scripts\new-git-activity-report.ps1` - Builds report metrics and markdown; can optionally skip PDF generation with `-SkipPdf`.
- `scripts\convert-markdown-to-pdf.ps1` - Converts markdown to a styled PDF.

