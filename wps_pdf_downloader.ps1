# SPDX-License-Identifier: MIT
# Copyright (c) 2025 Steve Chang

# wps_pdf_downloader.ps1
# Double-click (Run with PowerShell) → prompts for URL and output folder.
# Rules:
#   - If URL is blank → exit immediately.
#   - If folder is blank → use current folder (no new subfolder).
#   - If folder is an absolute path → use it as-is; otherwise create/use it under the script folder.
# Notes:
#   - Logs are saved to the chosen output folder ($OutDir) with datetime in the filename.
#   - Continuous mode can be enabled with the $EnableContinuousMode flag.
#   - File types are controlled by $AllowedExtensions (default: .pdf, .zip).

# --------------------------------------
# Settings
# --------------------------------------
# File name collision policy: Overwrite | Rename | Skip
$NameCollisionPolicy   = "Overwrite"

# Flag: set $true to enable continuous mode, $false to run only once
$EnableContinuousMode  = $true

# Also download these file extensions (edit this list to support more types)
$AllowedExtensions = @('.pdf', '.zip')

# Max retry attempts for each file download (used by Invoke-DownloadWithRetry)
$MaxDownloadRetries = 3

# --------------------------------------
# Helpers
# --------------------------------------
function Invoke-DownloaderRound {
    # Returns: [pscustomobject] @{ ExitCode=0|2|3|100; Ok=<int>; Fail=<int> }
    # ExitCode:
    #   0   = executed successfully (Fail may be >0)
    #   2   = fatal: page fetch failed
    #   3   = fatal: no matching links found (.pdf/.zip)
    #   100 = user aborted (blank URL)

    # Prompt URL
    $Url = Read-Host "Page URL (required; leave blank to exit)"
    if ([string]::IsNullOrWhiteSpace($Url)) {
        Write-Warning "No URL provided. Exiting."
        return [pscustomobject]@{ ExitCode=100; Ok=0; Fail=0 }
    }

    # Prompt output folder
    $FolderName = Read-Host "Output folder name (blank = current folder; absolute path OK)"
    $Root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

    # Expand environment variables like %USERPROFILE%
    if (-not [string]::IsNullOrWhiteSpace($FolderName)) {
        $FolderName = [Environment]::ExpandEnvironmentVariables($FolderName)
    }

    if ([string]::IsNullOrWhiteSpace($FolderName)) {
        $OutDir = $Root
    }
    elseif ([IO.Path]::IsPathRooted($FolderName)) {
        $OutDir = $FolderName
        if (-not (Test-Path -LiteralPath $OutDir)) {
            New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
        }
    }
    else {
        $OutDir = Join-Path -Path $Root -ChildPath $FolderName
        if (-not (Test-Path -LiteralPath $OutDir)) {
            New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
        }
    }

    # Start transcript
    $LogFile = Join-Path $OutDir ("log_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $TranscriptStarted = $false
    try {
        # NOTE: Transcript captures console I/O; avoid pasting credentials into prompts.
        Start-Transcript -Path $LogFile -Append -IncludeInvocationHeader | Out-Null
        $TranscriptStarted = $true
    } catch {
        Write-Warning "Failed to start transcript: $($_.Exception.Message)"
    }

    try {
        Write-Host ""
        Write-Host "Target URL : $Url"
        Write-Host "Output DIR : $OutDir"
        Write-Host "Log File   : $LogFile"
        Write-Host ""

        # Fetch page
        $resp = $null
        try { $resp = Invoke-WebRequest -Uri $Url -TimeoutSec 60 } catch {
            try { $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 60 } catch {
                Write-Error ("Failed to fetch page: {0}`n{1}" -f $Url, $_.Exception.Message)
                return [pscustomobject]@{ ExitCode=2; Ok=0; Fail=0 }
            }
        }

        # Extract links
        $hrefs = @()
        if ($resp.Links) { $hrefs += ($resp.Links | ForEach-Object { $_.href }) }
        if ($resp.Content) {
            $hrefMatches = Select-String -InputObject $resp.Content -Pattern 'href="([^"]+)"' -AllMatches
            if ($hrefMatches) { $hrefs += ($hrefMatches.Matches | ForEach-Object { $_.Groups[1].Value }) }
        }

        # Normalize & keep only allowed extensions
        $BaseUri = [System.Uri]$Url
        $urls = $hrefs |
            Where-Object { $_ } |
            ForEach-Object {
                $h = $_
                $h = [System.Net.WebUtility]::HtmlDecode($h).Trim()
                $h = $h.Trim('"', "'", ')')
                $h = ($h -replace '(%22|%27|%5C|&quot;|&#34;|&#39;)+$', '')
                $abs = $null

                if     ($h -match '^https?://') { $abs = $h }
                elseif ($h -like '//*')         { $abs = 'https:' + $h }
                else   { try { $abs = [System.Uri]::new($BaseUri, $h).AbsoluteUri } catch { $abs = $null } }

                if ($abs) {
                    $head = ($abs -split '#')[0]
                    $noq  = ($head -split '\?')[0]
                    $ext  = [IO.Path]::GetExtension($noq)

                    if ($AllowedExtensions -contains $ext.ToLowerInvariant()) { $abs }
                }
            } |
            Sort-Object -Unique

        if (-not $urls -or $urls.Count -eq 0) {
            Write-Warning "No matching links found (.pdf/.zip)."
            return [pscustomobject]@{ ExitCode=3; Ok=0; Fail=0 }
        }

        # Show counts by type
        $pdfCount = ($urls | Where-Object { [IO.Path]::GetExtension(($_ -split '\?')[0]).ToLowerInvariant() -eq '.pdf' }).Count
        $zipCount = ($urls | Where-Object { [IO.Path]::GetExtension(($_ -split '\?')[0]).ToLowerInvariant() -eq '.zip' }).Count
        Write-Host ("Found {0} link(s): PDF={1}, ZIP={2}" -f $urls.Count, $pdfCount, $zipCount) -ForegroundColor Green

        # Download all matched files (.pdf/.zip)
        $i=0; $ok=0; $fail=0
        foreach ($u in $urls) {
            $i++
            $name = [System.IO.Path]::GetFileName((($u -split '\?')[0]))
            if (-not $name) {
                # Pick extension based on URL
                $extFromUrl = [IO.Path]::GetExtension((($u -split '\?')[0]))
                if (-not $extFromUrl) { $extFromUrl = '.pdf' }  # fallback
                $name = "file_$i$extFromUrl"
            }

            $outPath = Resolve-DestinationPath -Dir $OutDir -FileName $name -Policy $NameCollisionPolicy
            if (-not $outPath) {
                # Skip strategy: still show which URL was associated
                Write-Host ("[{0}/{1}] SKIP (exists): {2}" -f $i, $urls.Count, $name) -ForegroundColor Yellow
                Write-Host ("     URL : {0}" -f $u)
                continue
            }

            $displayName = [IO.Path]::GetFileName($outPath)

            # Print filename + URL so Transcript records both
            Write-Host ("[{0}/{1}] {2}" -f $i, $urls.Count, $displayName)
            Write-Host ("     URL : {0}" -f $u)

            if (Invoke-DownloadWithRetry -Url $u -Dest $outPath -MaxRetries $MaxDownloadRetries) {
                $ok++
                Write-Host -NoNewline "     -> Status: "
                Write-Host "OK" -ForegroundColor Green
            } else {
                $fail++
                Write-Host -NoNewline "     -> Status: "
                Write-Host "Failed" -ForegroundColor Red
            }
        }

        Write-Host ""
        Write-Host ("Completed. Success: {0}  Failed: {1}" -f $ok, $fail)
        Write-Host ("Saved to: {0}" -f $OutDir)

        return [pscustomobject]@{ ExitCode=0; Ok=$ok; Fail=$fail }
    }
    finally {
        if ($TranscriptStarted) {
            try {
                Stop-Transcript | Out-Null
                Write-Host "Execution log saved: $LogFile"
            } catch {
                Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
            }
        }
    }
}

function Invoke-DownloadWithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$Dest,
        [int]$MaxRetries = 3,
        [hashtable]$Headers,
        [string]$UserAgent
    )

    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        $attempt++

        try {
            # $Headers / $UserAgent can be used to satisfy basic UA/Referer checks on some CDNs.
            Invoke-WebRequest `
                -Uri $Url `
                -OutFile $Dest `
                -Headers $Headers `
                -UserAgent $UserAgent `
                -MaximumRedirection 5 `
                -UseBasicParsing `
                -TimeoutSec 120

            if ((Test-Path -LiteralPath $Dest) -and ((Get-Item -LiteralPath $Dest).Length -gt 0)) {
                return $true
            } else {
                throw "Empty file"
            }
        }
        catch {
            # Extract HTTP status if available
            $statusCode = $null
            $statusText = $null
            try {
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode.value__
                    $statusText = [string]$_.Exception.Response.StatusDescription
                }
            } catch {}

            # Decide whether to retry based on status code
            $shouldRetry = $true
            if ($statusCode) {
                switch ($statusCode) {
                    400 { $shouldRetry = $false } # bad request – likely permanent
                    401 { $shouldRetry = $false } # unauthorized
                    403 { $shouldRetry = $true  } # forbidden – often referer/UA/CDN; retry after backoff
                    404 { $shouldRetry = $false } # not found – usually permanent
                    408 { $shouldRetry = $true  } # request timeout
                    {$_ -ge 429} { $shouldRetry = $true } # rate limit / server errors
                }
            }

            if ($attempt -lt $MaxRetries -and $shouldRetry) {
                # Exponential backoff with jitter (max 30s)
                $delay = [Math]::Min([int][Math]::Pow(2,$attempt) + (Get-Random -Min 0 -Max 3), 30)
                Write-Warning ("Retry {0}/{1}: {2}  {3}{4}" -f $attempt, $MaxRetries, $Url,
                    $(if ($statusCode) { "HTTP $statusCode " } else { "" }),
                    $(if ($statusText) { "($statusText)" } else { "" }) )
                Start-Sleep -Seconds $delay
            } else {
                # Final failure: print detailed reason
                $msg = if ($statusCode) {
                    "Failed: $Url  HTTP $statusCode" + $(if ($statusText) { " ($statusText)" } else { "" })
                } else {
                    "Failed: $Url  " + $_.Exception.Message
                }
                Write-Error $msg
                return $false
            }
        }
    }
}

function Resolve-DestinationPath {
    param(
        [Parameter(Mandatory=$true)][string]$Dir,
        [Parameter(Mandatory=$true)][string]$FileName,
        [ValidateSet('Overwrite','Rename','Skip')]
        [string]$Policy = 'Rename'
    )

    $target = Join-Path $Dir $FileName

    switch ($Policy) {
        'Overwrite' { return $target }

        'Skip' {
            if (Test-Path -LiteralPath $target) {
                Write-Host "Skip existing: $FileName"
                return $null
            }
            return $target
        }

        'Rename' {
            if (-not (Test-Path -LiteralPath $target)) { return $target }
            $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
            $ext  = [IO.Path]::GetExtension($FileName)
            for ($i = 2; $true; $i++) {
                $candidate = Join-Path $Dir ("{0} ({1}){2}" -f $base, $i, $ext)
                if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
            }
        }
    }
}
# --------------------------------------
# End of Helpers
# --------------------------------------

# Ensure TLS 1.2+ on legacy hosts
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} catch {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

# --------------------------------------
# Banner
# --------------------------------------
Write-Host ""
Write-Host "=== Windows PowerShell PDF Downloader v1.0 ===" -ForegroundColor Cyan
Write-Host ""
$Owner     = "Steve Chang"
$StartYear = 2025
$License   = "MIT License"
$currYear  = [int](Get-Date -Format yyyy)
$yearRange = if ($currYear -gt $StartYear) { "$StartYear-$currYear" } else { "$StartYear" }
Write-Host "(c) $yearRange $Owner - $License" -ForegroundColor DarkGray
Write-Host ""

# --------------------------------------
# Main Control
# --------------------------------------
if ($EnableContinuousMode) {
    while ($true) {
        $round = Invoke-DownloaderRound
        if ($round.ExitCode -eq 100) {
            break   # user aborted with blank URL
        }
        elseif ($round.ExitCode -eq 0 -and $round.Fail -eq 0) {
            Write-Host ""
            Write-Host "Round completed successfully. Starting the next round..." -ForegroundColor Yellow
            continue
        }
        else {
            Write-Host ""
            Write-Host "Round finished with errors. Please check the log for details." -ForegroundColor DarkYellow
            break
        }
    }
} else {
    # Single round only
    [void](Invoke-DownloaderRound)
}
