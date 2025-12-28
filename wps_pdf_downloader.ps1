# SPDX-License-Identifier: MIT
# Copyright (c) 2025 Steve Chang

param(
    [Parameter(Mandatory=$false)][Alias("u")][string]$UrlParam,
    [Parameter(Mandatory=$false)][Alias("o")][string]$OutDirParam,
    [Parameter(Mandatory=$false)][Alias("m")][ValidateSet("B", "S")][string]$ModeParam = "B",
    [Parameter(Mandatory=$false)][Alias("i")][switch]$InteractiveMode,
    [Parameter(Mandatory=$false)][Alias("c")][switch]$ContinuousMode
)

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
$NameCollisionPolicy = "Overwrite"

# Flag: set $true to enable continuous mode, $false to run only once
$EnableContinuousMode = if ($ContinuousMode) { $true } else { $false }

# If -c is used, we also force InteractiveMode to true
if ($ContinuousMode) {
    $InteractiveMode = $true
    Write-Host " [Info] Continuous Mode (-c) detected: Forcing Interactive Mode ON." -ForegroundColor Yellow
}

# Flag: set $true to save the original webpage HTML source
$SavePageSource = $true

# Also download these file extensions (edit this list to support more types)
$AllowedExtensions = @('.pdf', '.zip')

# Max retry attempts for each file download (used by Invoke-DownloadWithRetry)
$MaxDownloadRetries = 3

# --------------------------------------
# Environment Setup
# --------------------------------------
# Ensure terminal correctly renders special symbols like ® and Unicode characters
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --------------------------------------
# Helpers
# --------------------------------------

function Get-BrowserPath {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
    return $null
}

function Invoke-DownloaderRound {
    # Returns: [pscustomobject] @{ ExitCode=0|2|3|100; Ok=<int>; Fail=<int> }
    # ExitCode:
    #   0   = executed successfully (Fail may be >0)
    #   2   = fatal: page fetch failed
    #   3   = fatal: no matching links found (.pdf/.zip)
    #   100 = user aborted (blank URL)

    # 1. Handle URL
    $Url = if ($InteractiveMode) {
        Read-Host "Page URL (required; leave blank to exit)"
    } else { $UrlParam }
    if ([string]::IsNullOrWhiteSpace($Url)) {
        Write-Warning "No URL provided. Exiting."
        return [pscustomobject]@{ ExitCode=100; Ok=0; Fail=0 }
    }

    # 2. Handle Fetch Method (B=Browser, S=Static)
    $UseBrowser = $true
    $Mode = if ($InteractiveMode) {
        $input = Read-Host "Fetch Method? [B]rowser or [S]tatic"
        if ([string]::IsNullOrWhiteSpace($input)) { "" } else { $input.ToUpper() }
    } else {
        if ([string]::IsNullOrWhiteSpace($ModeParam)) { "" } else { $ModeParam.ToUpper() }
    }

    # Strict Validation: Terminate if input is neither B nor S
    if ($Mode -notin @("B", "S")) {
        Write-Error "[Error] Invalid Fetch Method: '$Mode'. Only 'B' or 'S' are allowed."
        return [pscustomobject]@{ ExitCode=100; Ok=0; Fail=0 }
    }

    if ($Mode -eq "S") { $UseBrowser = $false }

    # 3. Handle Output Directory
    $FolderName = if ($InteractiveMode) {
        Read-Host "Output folder name (blank = current folder; absolute path OK)"
    } else { $OutDirParam }

    $Root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

    # Support system environment variables (e.g., %USERPROFILE%\Downloads)
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

        $html = ""

        if ($UseBrowser) {
            # --- Option A: Browser Rendering ---
            # Use browser in Headless mode to render the web page
            $browser = Get-BrowserPath
            if (-not $browser) {
                # Write-Error "[Error] Browser not found. Headless rendering aborted."
                # return [pscustomobject]@{ ExitCode=2; Ok=0; Fail=0 }
                Write-Error "[Error] Browser not found. Falling back to Static Fetch."
                $UseBrowser = $false # Fallback if browser is missing
            } else {
                # Create temporary user profile for headless browser instance
                $tmpProfile = Join-Path ([IO.Path]::GetTempPath()) ("headless-localizer-" + [guid]::NewGuid())
                New-Item -ItemType Directory -Path $tmpProfile -Force | Out-Null

                try {
                    Write-Host "[Info] Rendering Web Page via Browser..." -ForegroundColor Yellow
                    $args = "--headless --disable-gpu --no-sandbox --user-data-dir=""$tmpProfile"" --dump-dom ""$Url"""
                    $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                        FileName = $browser; Arguments = $args; RedirectStandardOutput = $true; UseShellExecute = $false; CreateNoWindow = $true
                    }
                    $p = [System.Diagnostics.Process]::Start($psi)
                    $startTime = Get-Date

                    # Read rendered HTML output from the browser. ReadToEnd() is
                    # synchronous. For extremely large DOMs (>64KB), if the browser
                    # process blocks, this line might cause a deadlock.
                    $html = $p.StandardOutput.ReadToEnd()
                    # Write-Host -NoNewline $html
                    Write-Host ""
                    Write-Host "[Info] Total Characters : $($html.Length)"
                    Write-Host "[Info] Elapsed Time     : $(((Get-Date) - $startTime).TotalSeconds) seconds"
                    Write-Host ""
                } finally {
                    # Ensure browser process is terminated and cleanup temp profile
                    if ($null -ne $p -and -not $p.HasExited) { try { $p.Kill() } catch { } }
                    Remove-Item $tmpProfile -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }

        if (-not $UseBrowser) {
            # --- Option B: Static Fetch (Invoke-WebRequest) ---
            try {
                Write-Host "[Info] Fetching Page via Invoke-WebRequest..." -ForegroundColor Yellow
                $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 60
                $html = $resp.Content
            } catch {
                Write-Error ("Failed to fetch page: {0}`n{1}" -f $Url, $_.Exception.Message)
                return [pscustomobject]@{ ExitCode=2; Ok=0; Fail=0 }
            }
        }

        # Archive the original HTML source if enabled
        if ($Global:SavePageSource -and $html) {
            # Extract Page Title using Regex
            $TitleMatch = [regex]::Match($html, '(?s)<title>(.*?)</title>', "IgnoreCase")
            $SafeTitle = "source" # Default fallback

            if ($TitleMatch.Success) {
                # Decode HTML entities (e.g., &reg; -> ®) and trim whitespaces
                $RawTitle = [System.Net.WebUtility]::HtmlDecode($TitleMatch.Groups[1].Value.Trim())
                # Replace illegal filename characters with underscores
                $SafeTitle = $RawTitle -replace '[\\/:*?"<>|]', '_'
            }

            # Create a unique filename with timestamp
            $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

            # Create resources folder for localizing images
            $FilesFolder = "${SafeTitle}_${Timestamp}_files"
            $FilesPath = Join-Path $OutDir $FilesFolder
            if (-not (Test-Path $FilesPath)) { New-Item $FilesPath -ItemType Directory -Force | Out-Null }

            Write-Host "[Info] Localizing images into: $FilesFolder" -ForegroundColor Cyan
            Write-Host ""
            $imgMatches = [regex]::Matches($html, '<img\s+[^>]*src="([^"]+)"')
            $BaseUri = [System.Uri]$Url
            $imgCount = 0

            foreach ($match in $imgMatches) {
                $rawSrc = $match.Groups[1].Value
                try {
                    $absImgUrl = [System.Uri]::new($BaseUri, $rawSrc).AbsoluteUri
                    $imgName = [IO.Path]::GetFileName(($absImgUrl -split '\?')[0])
                    if (-not $imgName) { $imgName = "image_$imgCount.png" }

                    $localImgPath = Join-Path $FilesPath $imgName
                    $relativeImgPath = "./$FilesFolder/$imgName"

                    Write-Host ("[Info] Downloading image[{0}]: {1}" -f $imgCount, $absImgUrl) # -ForegroundColor DarkGray

                    if (Invoke-DownloadWithRetry -Url $absImgUrl -Dest $localImgPath -MaxRetries $Global:MaxDownloadRetries) {
                        # Replace remote URL with local relative path in HTML
                        $html = $html.Replace($rawSrc, $relativeImgPath)
                        $imgCount++
                    }
                } catch {}
            }
            Write-Host ""
            Write-Host ("[Info] Found {0} image(s)." -f $imgCount) -ForegroundColor Green
            Write-Host ""

            $HtmlName = "{0}_{1}.html" -f $SafeTitle, $Timestamp
            $HtmlPath = Join-Path $OutDir $HtmlName

            try {
                # Save the localized HTML source using UTF8 encoding
                $html | Out-File -FilePath $HtmlPath -Encoding utf8
                Write-Host "[Info] Web page saved to: $HtmlName" # -ForegroundColor Cyan
            } catch {
                Write-Warning "Failed to save web page source: $($_.Exception.Message)"
            }
        }

        # --- Extract links (Hybrid Method) ---
        $hrefs = @()
        # Source A: From Static fetch (Invoke-WebRequest) links object
        if ($null -ne $resp -and $resp.Links) {
            $hrefs += ($resp.Links | ForEach-Object { $_.href })
        }
        # Source B: From Dynamic render (Headless Browser) raw HTML via Regex
        if ($html) {
            $hrefMatches = Select-String -InputObject $html -Pattern 'href="([^"]+)"' -AllMatches
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

        # Show download counts categorized by extension
        $pdfCount = ($urls | Where-Object { [IO.Path]::GetExtension(($_ -split '\?')[0]).ToLowerInvariant() -eq '.pdf' }).Count
        $zipCount = ($urls | Where-Object { [IO.Path]::GetExtension(($_ -split '\?')[0]).ToLowerInvariant() -eq '.zip' }).Count
        Write-Host ""
        Write-Host ("[Info] Found {0} link(s). PDF={1} ZIP={2}" -f $urls.Count, $pdfCount, $zipCount) -ForegroundColor Green
        Write-Host ""

        # Download all matched files with collision handling
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
        # Finalize transcript log
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
            break   # User input blank URL to stop
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
