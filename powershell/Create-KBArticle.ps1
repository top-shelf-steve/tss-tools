<#
.SYNOPSIS
    Interactively creates a KB article in markdown format.
.DESCRIPTION
    Prompts for title, description, steps, and optional notes.
    Scans the current directory for images and lets you attach them to steps.
    Outputs a clean markdown file named KB-{Title}.md.
.EXAMPLE
    cd C:\KBDrafts\password-reset
    .\Create-KBArticle.ps1
#>

function Get-MultiLineInput {
    param([string]$Prompt)
    Write-Host "$Prompt" -ForegroundColor Cyan
    Write-Host "(Enter a blank line when finished)" -ForegroundColor DarkGray
    $lines = @()
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line) -and $lines.Count -gt 0) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $lines += $line
    }
    return $lines -join "`n"
}

function Get-DirectoryImages {
    $extensions = @('*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp', '*.svg')
    $images = @()
    foreach ($ext in $extensions) {
        $images += Get-ChildItem -Path . -Filter $ext -File -ErrorAction SilentlyContinue
    }
    return $images | Sort-Object Name
}

function Select-Image {
    param([array]$Images)
    if ($Images.Count -eq 0) { return $null }

    Write-Host "`nImages found in current directory:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Images.Count; $i++) {
        Write-Host "  [$($i + 1)] $($Images[$i].Name)" -ForegroundColor White
    }
    Write-Host "  [0] None" -ForegroundColor DarkGray

    $selection = Read-Host "Attach an image to this step?"
    if ([string]::IsNullOrWhiteSpace($selection) -or $selection -eq '0') { return $null }

    $index = 0
    if ([int]::TryParse($selection, [ref]$index) -and $index -ge 1 -and $index -le $Images.Count) {
        return $Images[$index - 1]
    }

    Write-Host "Invalid selection, skipping." -ForegroundColor DarkGray
    return $null
}

# --- Main ---

Write-Host "`n=== KB Article Creator ===" -ForegroundColor Green
Write-Host ""

# Title
$title = Read-Host "Article Title"
while ([string]::IsNullOrWhiteSpace($title)) {
    $title = Read-Host "Title cannot be blank. Article Title"
}

# Description
Write-Host ""
$description = Get-MultiLineInput -Prompt "Description:"

# Find images once
$images = Get-DirectoryImages

# Steps
Write-Host ""
Write-Host "--- Steps ---" -ForegroundColor Green
$steps = @()
$stepNum = 1

while ($true) {
    Write-Host ""
    $stepText = Get-MultiLineInput -Prompt "Step ${stepNum}:"

    $stepImage = $null
    if ($images.Count -gt 0) {
        $stepImage = Select-Image -Images $images
    }

    $steps += [PSCustomObject]@{
        Text  = $stepText
        Image = $stepImage
    }

    Write-Host ""
    $more = Read-Host "Add another step? (y/n)"
    if ($more -notmatch '^y') { break }
    $stepNum++
}

# Notes
Write-Host ""
$notes = Read-Host "Notes (optional, press Enter to skip)"

# --- Build Markdown ---

$md = @()
$md += "# KB: $title"
$md += ""
$md += "## Description"
$md += ""
$md += $description
$md += ""
$md += "## Steps"
$md += ""

for ($i = 0; $i -lt $steps.Count; $i++) {
    $md += "### Step $($i + 1)"
    $md += ""
    $md += $steps[$i].Text
    $md += ""
    if ($null -ne $steps[$i].Image) {
        $imgName = $steps[$i].Image.Name
        $imgLabel = [System.IO.Path]::GetFileNameWithoutExtension($imgName)
        $md += "![$imgLabel](./$imgName)"
        $md += ""
    }
}

if (-not [string]::IsNullOrWhiteSpace($notes)) {
    $md += "## Notes"
    $md += ""
    $md += $notes
    $md += ""
}

# --- Write File ---

$safeName = $title -replace '[^\w\s-]', '' -replace '\s+', '-'
$filename = "KB-$safeName.md"
$outputPath = Join-Path -Path (Get-Location) -ChildPath $filename

($md -join "`n") | Out-File -FilePath $outputPath -Encoding utf8

Write-Host ""
Write-Host "KB article created: $outputPath" -ForegroundColor Green
