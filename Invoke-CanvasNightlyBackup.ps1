#Requires -Version 5.1
<#
.SYNOPSIS
    Nightly Canvas LMS backup - self-contained PowerShell, no NPM/Node/LLM required.

.DESCRIPTION
    This script performs a nightly backup workflow for Canvas content and gradebooks.
    Uses extensive parallel background jobs to achieve maximum throughput.

    High-level flow:
      1. Load encrypted configuration from C:\Scripts\canvas-backup-config.clixml
      2. Resolve terms and enumerate active courses
      3. Bulk-cache today's existing S3 backups to memory (High-Speed Deduplication)
      4. Back up course content packages (.imscc) - (Configurable Parallel Background Jobs)
      5. Back up gradebook CSVs - (Configurable Parallel Background Jobs)
      6. Back up course Pages (JSON with HTML body) - (Configurable Parallel Background Jobs)
      7. Back up course Files (.zip archive of raw files) - (Configurable Parallel Background Jobs)
         * Includes an automated HDD cache safeguard to ensure local temp files never exceed 100GB.
         * Features Bin Packing: Auto-splits courses > 20GB into multiple 10GB zip parts to protect RAM/HDD.
         * Uses native .NET ZipFile compression to bypass PowerShell's 2GB file-locking bugs.
         * Hardened Try/Catch/Finally to skip 404 ghost files and guarantee lock releases.
      8. Optionally recurse through all Canvas sub-accounts for persistent backups
      9. Verify S3 coverage, upload run logs, and perform final HDD sweep

    Configuration Setup (Run Once):
      Run the following block manually in PowerShell as the SAME Windows user 
      account that will run the scheduled task. This creates the encrypted 
      credentials file at C:\Scripts\canvas-backup-config.clixml:
      
      $configPath = 'C:\Scripts\canvas-backup-config.clixml'
      $config = [pscustomobject]@{
          CanvasBaseUrl   = Read-Host 'Canvas Base URL'
          S3Bucket        = Read-Host 'S3 Bucket'
          AwsRegion       = Read-Host 'AWS Region'
          RootAccountId   = Read-Host 'Root Account ID (Default is 1)'
          CanvasApiToken  = Read-Host 'Canvas API Token' -AsSecureString
      }
      $config | Export-Clixml -Path $configPath

    Configure AWS Authentication & Optimization (Run Once):
      Ensure the Windows user account running the script is authenticated to AWS, 
      and configure the concurrency settings to support the script's parallel background jobs.
      Run this in your terminal:

      aws configure set aws_access_key_id YOUR_KEY
      aws configure set aws_secret_access_key YOUR_SECRET
      aws configure set default.region us-east-1
      aws configure set default.output json
      # Required for parallel pushes:
      aws configure set default.s3.max_concurrent_requests 50
      aws configure set default.s3.max_queue_size 10000

    Hard dependencies:
      - PowerShell 5.1+ (ships with Windows; no install needed)
      - AWS CLI v2 (aws.exe)
      - Network access to the Canvas REST API and AWS S3

.PARAMETER CanvasBaseUrl
    Canvas instance URL, e.g. https://your-school.instructure.com

.PARAMETER CanvasApiToken
    Canvas API bearer token with admin read + export scope.

.PARAMETER S3Bucket
    S3 bucket name for all backup writes.

.PARAMETER AwsRegion
    AWS region for the S3 bucket. Defaults to us-east-1.

.PARAMETER RootAccountId
    Canvas root account ID. Default: 1.

.PARAMETER TempDir
    Local scratch directory for in-flight export downloads.

.PARAMETER LogDir
    Directory for transcript log files.

.PARAMETER ConcurrentJobs
    Number of parallel background jobs to run for data extraction.

.PARAMETER CourseId
    Optional single-course mode.

.PARAMETER MaxCourses
    Optional test-mode limiter.

.PARAMETER SkipSubAccounts
    Skip the recursive sub-account backup pass.

.PARAMETER SkipContent
    Skip course content (.imscc) export backups.

.PARAMETER SkipPages
    Skip course Pages JSON backups.

.PARAMETER SkipFiles
    Skip course raw Files backups.

.PARAMETER SkipGradebooks
    Skip gradebook CSV backups.

.PARAMETER WhatIfMode
    Dry-run switch. No Canvas exports, S3 uploads, S3 copies, or S3 deletes occur.

.PARAMETER Force
    Bypass same-day deduplication and always submit a fresh Canvas export.
#>

param(
    [string]$CanvasBaseUrl,
    [string]$CanvasApiToken,
    [string]$S3Bucket,
    [string]$AwsRegion,
    [string]$RootAccountId,
    [string]$TempDir,
    [string]$LogDir,
    [int]$PollTimeoutMins = 60,
    [int]$PollIntervalSecs = 3,
    [int]$ConcurrentJobs = 5,
    [string]$CourseId,
    [int]$MaxCourses = 0,
    [switch]$SkipSubAccounts,
    [switch]$SkipContent,
    [switch]$SkipPages,
    [switch]$SkipFiles,
    [switch]$SkipGradebooks,
    [switch]$WhatIfMode,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure a minimum of 1 concurrent job to prevent math errors
if ($ConcurrentJobs -lt 1) { $ConcurrentJobs = 1 }

# ---------------------------------------------------------------------------
# Fixed working paths
# ---------------------------------------------------------------------------
$script:BaseScriptDir   = 'C:\Scripts'
$script:ConfigPath      = 'C:\Scripts\canvas-backup-config.clixml'
$script:DefaultLogDir   = 'C:\Scripts\Logs'
$script:DefaultTempDir  = 'C:\Scripts\Temp'

if (-not $LogDir)  { $LogDir  = $script:DefaultLogDir }
if (-not $TempDir) { $TempDir = $script:DefaultTempDir }

# ============================================================
# Encrypted configuration loading
# ============================================================
function Import-CanvasBackupConfig {
    param([string]$Path = $script:ConfigPath)
    if (-not (Test-Path $Path)) { return $null }
    Import-Clixml -Path $Path
}

function ConvertFrom-SecureStringToPlainText {
    param([Security.SecureString]$SecureString)
    if ($null -eq $SecureString) { return $null }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

# ============================================================
# Console logging helpers
# ============================================================
function Write-Info   ([string]$msg) { Write-Host "[INFO ] $msg"  -ForegroundColor Cyan }
function Write-Ok     ([string]$msg) { Write-Host "[ OK  ] $msg"  -ForegroundColor Green }
function Write-Warn   ([string]$msg) { Write-Host "[WARN ] $msg"  -ForegroundColor Yellow }
function Write-Fail   ([string]$msg) { Write-Host "[FAIL ] $msg"  -ForegroundColor Red }
function Write-DryRun ([string]$msg) { Write-Host "[DRYRUN] $msg" -ForegroundColor Magenta }

# ============================================================
# Persistent Temp Cleanup Helper (Solves Antivirus File Locking)
# ============================================================
function Remove-TempPath {
    param([string]$Path, [switch]$Recurse)
    if (-not (Test-Path $Path)) { return }
    Start-Sleep -Seconds 1 
    $delRetries = 90
    while ((Test-Path $Path) -and $delRetries -gt 0) {
        if ($Recurse) { Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue }
        else { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
        if (Test-Path $Path) { Start-Sleep -Seconds 2; $delRetries-- }
    }
    if (Test-Path $Path) { Write-Warn "Could not delete local path after 3 minutes! File severely locked: $Path" }
}

# ============================================================
# Canvas API helpers
# ============================================================
function Get-JsonValue {
    param($Obj, [string]$Property)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [hashtable]) { return $Obj[$Property] }
    if ($Obj.PSObject.Properties.Match($Property).Count -gt 0) { return $Obj.$Property }
    return $null
}

function Invoke-CanvasApiPage {
    param([string]$Url)
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $resp = Invoke-WebRequest -Uri $Url -Method Get `
                -Headers @{ Authorization = "Bearer $CanvasApiToken" } `
                -UseBasicParsing -TimeoutSec 60
            return $resp
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response
            $code = $null
            if ($response -and $response.StatusCode) { $code = [int]$response.StatusCode }

            if ($code -eq 429) {
                $ra = 10
                $raHeader = $response.Headers['Retry-After']
                if ($raHeader) { try { $ra = [int]$raHeader } catch {} }
                Write-Warn "Rate-limited - waiting $ra s"
                Start-Sleep -Seconds $ra
            }
            elseif ($attempt -lt 3) {
                Start-Sleep -Seconds (5 * $attempt)
            }
            else { throw }
        }
    }
}

function Invoke-CanvasGet {
    param([string]$Endpoint, [hashtable]$Query = @{})
    $base = $CanvasBaseUrl.TrimEnd('/')
    $qs = @()
    $hasPerPage = $false
    foreach ($kv in $Query.GetEnumerator()) {
        if ($kv.Key -eq 'per_page') { $hasPerPage = $true }
        $qs += "$([Uri]::EscapeDataString($kv.Key))=$([Uri]::EscapeDataString([string]$kv.Value))"
    }
    if (-not $hasPerPage) { $qs += 'per_page=100' }

    $url = "$base/api/v1${Endpoint}?" + ($qs -join '&')
    $results = @()

    do {
        $resp = Invoke-CanvasApiPage -Url $url
        $page = $resp.Content | ConvertFrom-Json
        if ($null -ne $page) { $results += @($page) }
        $url = $null
        $linkHeader = $resp.Headers['Link']
        if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) { $url = $Matches[1] }
    } while ($url)

    return $results
}

function Get-EnrollmentTerms {
    param([string]$AccountId)
    $base = $CanvasBaseUrl.TrimEnd('/')
    $url = "$base/api/v1/accounts/${AccountId}/terms?per_page=100"
    $results = @()

    do {
        $resp = Invoke-CanvasApiPage -Url $url
        $page = $resp.Content | ConvertFrom-Json
        $terms = Get-JsonValue -Obj $page -Property 'enrollment_terms'
        if ($terms) { $results += @($terms) }
        $url = $null
        $linkHeader = $resp.Headers['Link']
        if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) { $url = $Matches[1] }
    } while ($url)

    return $results
}

function Invoke-CanvasPost {
    param([string]$Endpoint, [hashtable]$Body = @{})
    $uri = "$($CanvasBaseUrl.TrimEnd('/'))/api/v1${Endpoint}"
    $json = $Body | ConvertTo-Json -Compress
    $resp = Invoke-WebRequest -Uri $uri -Method Post `
        -Headers @{ Authorization = "Bearer $CanvasApiToken" } `
        -ContentType 'application/json' -Body $json -UseBasicParsing -TimeoutSec 60
    $resp.Content | ConvertFrom-Json
}

function Get-CanvasCourseById {
    param([string]$CourseId)
    $base = $CanvasBaseUrl.TrimEnd('/')
    $url = "$base/api/v1/courses/${CourseId}?include[]=term&include[]=account_name"
    Write-Info "Fetching single course: $CourseId"
    try {
        $resp = Invoke-CanvasApiPage -Url $url
        return ($resp.Content | ConvertFrom-Json)
    }
    catch [System.Net.WebException] {
        $response = $_.Exception.Response
        if ($response -and $response.StatusCode -eq 404) {
            Write-Warn "CourseId $CourseId was not found or is not accessible"
            return $null
        }
        throw
    }
}

function Get-SubAccounts {
    param([string]$AccountId)
    @(Invoke-CanvasGet "/accounts/${AccountId}/sub_accounts")
}

function Get-AllSubAccountsRecursive {
    param([string]$RootId)
    $all = New-Object System.Collections.Generic.List[object]
    $queue = New-Object System.Collections.Generic.Queue[object]
    $seen = @{}

    foreach ($sa in @(Get-SubAccounts -AccountId $RootId)) {
        if (-not $seen.ContainsKey("$($sa.id)")) {
            $seen["$($sa.id)"] = $true
            $queue.Enqueue($sa)
        }
    }

    while ($queue.Count -gt 0) {
        $acct = $queue.Dequeue()
        $all.Add($acct)
        foreach ($child in @(Get-SubAccounts -AccountId $acct.id)) {
            if (-not $seen.ContainsKey("$($child.id)")) {
                $seen["$($child.id)"] = $true
                $queue.Enqueue($child)
            }
        }
    }
    @($all)
}

# ============================================================
# S3 helpers
# ============================================================
function Test-S3KeyExists {
    param([string]$Key)
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    
    $json = & aws s3api list-objects-v2 --bucket $S3Bucket --prefix $Key --max-items 1 --output json --region $AwsRegion 2>$null
    $ErrorActionPreference = $oldEAP

    if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { return $false }
    $page = $json | ConvertFrom-Json
    if ($null -eq $page) { return $false }
    $hasContents = Get-JsonValue -Obj $page -Property 'Contents'
    if (-not $hasContents) { return $false }

    return @($page.Contents | Where-Object { $_.Key -eq $Key }).Count -gt 0
}

function Publish-FileToS3 {
    param([string]$LocalPath, [string]$Key, [string]$StorageClass = 'GLACIER_IR')
    if ($WhatIfMode) {
        Write-DryRun "Would upload '$LocalPath' to s3://$S3Bucket/$Key (storage-class=$StorageClass)"
        return
    }

    $maxRetries = 3
    $attempt = 0
    $success = $false

    while ($attempt -lt $maxRetries -and -not $success) {
        $attempt++
        & aws s3 cp $LocalPath "s3://$S3Bucket/$Key" --storage-class $StorageClass --region $AwsRegion --no-progress
        
        if ($LASTEXITCODE -eq 0) {
            $success = $true
        } else {
            if ($attempt -lt $maxRetries) {
                $sleepTime = $attempt * 10
                Write-Warn "S3 Upload dropped for $Key. Retrying in $sleepTime seconds... (Attempt $attempt of $maxRetries)"
                Start-Sleep -Seconds $sleepTime
            }
        }
    }

    if (-not $success) { throw "aws s3 cp failed for key: $Key after $maxRetries attempts." }
}

function Publish-TextToS3 {
    param([string]$Text, [string]$Key)
    $tmpFile = Join-Path $TempDir "s3-text-$([Guid]::NewGuid()).txt"
    try {
        [System.IO.File]::WriteAllText($tmpFile, $Text, [System.Text.Encoding]::UTF8)
        Publish-FileToS3 -LocalPath $tmpFile -Key $Key -StorageClass 'STANDARD'
    }
    finally {
        Remove-TempPath -Path $tmpFile
    }
}

function Invoke-S3CopyObject {
    param([string]$SourceKey, [string]$DestKey)
    if ($WhatIfMode) {
        Write-DryRun "Would copy s3://$S3Bucket/$SourceKey -> s3://$S3Bucket/$DestKey"
        return
    }
    & aws s3 cp "s3://$S3Bucket/$SourceKey" "s3://$S3Bucket/$DestKey" --metadata-directive COPY --storage-class GLACIER_IR --region $AwsRegion --no-progress
    if ($LASTEXITCODE -ne 0) { throw "aws s3 cp (copy) failed: $SourceKey -> $DestKey" }
}

function Get-S3ObjectsUnderPrefix {
    param([string]$Prefix)
    $results = @()
    $token = $null
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    do {
        $awsArgs = @('s3api', 'list-objects-v2', '--bucket', $S3Bucket, '--prefix', $Prefix, '--output', 'json', '--region', $AwsRegion)
        if ($token) { $awsArgs += @('--continuation-token', $token) }

        $json = & aws @awsArgs 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }

        $page = $json | ConvertFrom-Json
        if ($null -eq $page) { break }

        $hasContents = Get-JsonValue -Obj $page -Property 'Contents'
        if ($hasContents) {
            $results += @($page.Contents | Select-Object Key, LastModified)
        }

        $isTrunc = Get-JsonValue -Obj $page -Property 'IsTruncated'
        $nxtToken = Get-JsonValue -Obj $page -Property 'NextContinuationToken'

        if ($isTrunc -and $nxtToken) {
            $token = $page.NextContinuationToken
        } else { $token = $null }
    } while ($token)

    $ErrorActionPreference = $oldEAP
    if (-not $results) { return @() }
    return @($results | Sort-Object { [datetime]$_.LastModified } -Descending)
}

function Get-S3KeysUnderPrefix {
    param([string]$Prefix)
    $results = @()
    $token = $null
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    do {
        $awsArgs = @('s3api', 'list-objects-v2', '--bucket', $S3Bucket, '--prefix', $Prefix, '--output', 'json', '--region', $AwsRegion)
        if ($token) { $awsArgs += @('--continuation-token', $token) }

        $json = & aws @awsArgs 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }

        $page = $json | ConvertFrom-Json
        if ($null -eq $page) { break }

        $hasContents = Get-JsonValue -Obj $page -Property 'Contents'
        if ($hasContents) {
            $results += @($page.Contents | ForEach-Object { $_.Key })
        }

        $isTrunc = Get-JsonValue -Obj $page -Property 'IsTruncated'
        $nxtToken = Get-JsonValue -Obj $page -Property 'NextContinuationToken'

        if ($isTrunc -and $nxtToken) {
            $token = $page.NextContinuationToken
        } else { $token = $null }
    } while ($token)

    $ErrorActionPreference = $oldEAP
    return @($results)
}

function Invoke-S3Prune {
    param([string]$Prefix, [int]$KeepCount, [string]$Ext)
    $objects = @(Get-S3ObjectsUnderPrefix -Prefix $Prefix | Where-Object { $_.Key.EndsWith($Ext) })
    if ($objects.Count -le $KeepCount) { return }

    $toDelete = @($objects | Select-Object -Skip $KeepCount)
    if ($toDelete.Count -ge $objects.Count) { return }

    foreach ($obj in $toDelete) {
        if ($WhatIfMode) { Write-DryRun "Would delete s3://$S3Bucket/$($obj.Key)" }
        else { & aws s3api delete-object --bucket $S3Bucket --key $obj.Key --region $AwsRegion | Out-Null }
    }
}

function Invoke-TieredPromotion {
    param([string]$BaseKey, [string]$Today, [string]$Ext)
    $dailyKey = "$BaseKey/daily/$Today$Ext"
    $weeklyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/weekly/")
    $monthlyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/monthly/")

    $daysSinceWeekly = 999
    $daysSinceMonthly = 999

    if ($weeklyObjs.Count -gt 0) { $daysSinceWeekly = ((Get-Date) - [datetime]$weeklyObjs[0].LastModified).TotalDays }
    if ($monthlyObjs.Count -gt 0) { $daysSinceMonthly = ((Get-Date) - [datetime]$monthlyObjs[0].LastModified).TotalDays }

    $needsWeekly = $daysSinceWeekly -ge 7
    $needsMonthly = $daysSinceMonthly -ge 30

    if ($needsWeekly) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/weekly/$Today$Ext" }
    if ($needsWeekly -and $needsMonthly) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/monthly/$Today$Ext" }

    Invoke-S3Prune -Prefix "$BaseKey/daily/" -KeepCount 14 -Ext $Ext
    Invoke-S3Prune -Prefix "$BaseKey/weekly/" -KeepCount 5 -Ext $Ext
    Invoke-S3Prune -Prefix "$BaseKey/monthly/" -KeepCount 6 -Ext $Ext
}

# ============================================================
# Key generation helpers
# ============================================================
function ConvertTo-CanvasSlug {
    param([string]$Text)
    if (-not $Text) { return '' }
    (($Text.ToLower() -replace '[^a-z0-9]+', '-') -replace '^-+|-+$', '')
}

$script:GENERIC_ACCOUNTS = @(
    'academic-courses', 'courses', 'undergraduate', 'graduate',
    'academic', 'all-courses', 'default-term', 'root-account',
    'manually-created-courses'
)

function Get-TermParts {
    param([string]$TermName, [string]$StartAt)
    $season = $null
    $year = $null

    if ($TermName) {
        if ($TermName -match '\b(fall|spring|summer|winter)\b.*?(\d{4})') {
            $season = $Matches[1].ToLower(); $year = $Matches[2]
        }
        elseif ($TermName -match '(\d{4}).*?\b(fall|spring|summer|winter)\b') {
            $year = $Matches[1]; $season = $Matches[2].ToLower()
        }
        elseif ($TermName -match '(?i)jan(uary)?[\s\-]?(term|session)?|j[\s\.\-]?term') {
            $season = 'winter'
            if ($TermName -match '(\d{4})') { $year = $Matches[1] }
        }
    }

    if ($StartAt -and (-not $season -or -not $year)) {
        try {
            $dt = [datetime]$StartAt
            if (-not $year) { $year = $dt.Year.ToString() }
            if (-not $season) {
                $season = switch ($dt.Month) {
                    { $_ -in @(12,1) } { 'winter' }
                    { $_ -in 2..5 }    { 'spring' }
                    { $_ -in 6..8 }    { 'summer' }
                    default            { 'fall' }
                }
            }
        } catch {}
    }

    if (-not $year) {
        if ($TermName -and $TermName -match '(\d{4})') { $year = $Matches[1] }
        else { $year = (Get-Date).Year.ToString() }
    }
    if (-not $season) { $season = 'unknown' }

    [pscustomobject]@{ Year = $year; Semester = $season }
}

function Get-DeptSlug {
    param([string]$AccountSlug, [string]$CourseCode)
    if ($AccountSlug -and $AccountSlug -notin $script:GENERIC_ACCOUNTS) { return $AccountSlug }
    if ($CourseCode -and $CourseCode -match '^([A-Za-z]+)') { return $Matches[1].ToLower() }
    if ($AccountSlug) { return $AccountSlug }
    'other'
}

function Get-CourseS3BaseKey {
    param($Course, $TermParts, [string]$DeptSlug, [string]$BasePrefix = 'canvas-backups')
    $sis = Get-JsonValue -Obj $Course -Property 'sis_course_id'
    $identifier = if ($sis) { $sis } else { "$((ConvertTo-CanvasSlug (Get-JsonValue -Obj $Course -Property 'name')))-$($TermParts.Semester)-$($TermParts.Year)-$($Course.id)" }
    "$BasePrefix/$($TermParts.Year)/$($TermParts.Semester)/$DeptSlug/$identifier"
}

# ============================================================
# Modular Backup Routines (Used for Sub-Account Loop)
# ============================================================
function Backup-CourseContent {
    param($ExportData)
    Set-StrictMode -Off

    $courseId = $ExportData.CourseId
    $dailyKey = $ExportData.DailyContent
    $baseKey = $ExportData.BaseKey

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Content -> s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }
    if (-not $Force -and (Test-S3KeyExists -Key $dailyKey)) {
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Submitting content export..."
    try {
        $job = Invoke-CanvasPost "/courses/$courseId/content_exports" @{ export_type = 'common_cartridge'; skip_notifications = $true }
    } catch {
        Write-Fail "[$courseId] Submit failed: $_"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "Submit: $_" }
    }

    $deadline = (Get-Date).AddMinutes($PollTimeoutMins)
    $exp = $null

    do {
        Start-Sleep -Seconds $PollIntervalSecs
        if ((Get-Date) -gt $deadline) {
            Write-Fail "[$courseId] Timed out waiting for export"
            return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Poll timeout' }
        }
        try { $exp = (Invoke-CanvasGet "/courses/$courseId/content_exports/$($job.id)" | Select-Object -First 1) } catch { continue }
        if ($exp -and $exp.workflow_state -eq 'failed') {
            Write-Fail "[$courseId] Canvas export failed"
            return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Canvas export failed' }
        }
    } until ($exp -and $exp.workflow_state -eq 'exported')

    if (-not $exp.attachment -or -not $exp.attachment.url) { return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Missing URL' } }

    $downloadUrl = $exp.attachment.url
    $tempFile = Join-Path $TempDir "canvas-export-$courseId-$([Guid]::NewGuid()).imscc"

    try {
        Write-Info "[$courseId] Downloading content..."
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 600
        Publish-FileToS3 -LocalPath $tempFile -Key $dailyKey
        Invoke-TieredPromotion -BaseKey $baseKey -Today (Get-Date -Format 'yyyy-MM-dd') -Ext '.imscc'
        Write-Ok "[$courseId] $([Math]::Round((Get-Item $tempFile).Length / 1MB, 1)) MB -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{ CourseId = $courseId; Action = 'written'; S3Key = $dailyKey }
    } catch {
        Write-Fail "[$courseId] Upload/download error: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    } finally {
        Remove-TempPath -Path $tempFile
    }
}

function Backup-CoursePages {
    param($ExportData, [hashtable]$S3Cache = $null)
    Set-StrictMode -Off

    $courseId = $ExportData.CourseId
    $dailyKey = $ExportData.DailyPages
    $pagesBaseKey = "$($ExportData.BaseKey)-pages"

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Pages -> s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }

    $exists = $false
    if ($null -ne $S3Cache) { $exists = $S3Cache.ContainsKey($dailyKey) } else { $exists = Test-S3KeyExists -Key $dailyKey }
    if (-not $Force -and $exists) {
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Fetching pages data..."
    try {
        $pages = @(Invoke-CanvasGet "/courses/$courseId/pages" @{ 'include[]' = 'body' })
        if ($pages.Count -eq 0) { return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; SkipReason = 'no pages' } }

        $jsonText = $pages | ConvertTo-Json -Depth 10 -Compress
        Publish-TextToS3 -Text $jsonText -Key $dailyKey
        Invoke-TieredPromotion -BaseKey $pagesBaseKey -Today (Get-Date -Format 'yyyy-MM-dd') -Ext '.json'

        Write-Ok "[$courseId] Pages ($($pages.Count) items) -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{ CourseId = $courseId; Action = 'written'; S3Key = $dailyKey }
    }
    catch {
        Write-Fail "[$courseId] Pages fetch/upload failed: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }
}

function Backup-CourseFiles {
    param($ExportData, [hashtable]$S3Cache = $null)
    Set-StrictMode -Off

    $courseId = $ExportData.CourseId
    $dailyKey = $ExportData.DailyFiles
    $filesBaseKey = "$($ExportData.BaseKey)-files"

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Files Zip -> s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }

    $exists = $false
    if ($null -ne $S3Cache) { $exists = $S3Cache.ContainsKey($dailyKey) } else { $exists = Test-S3KeyExists -Key $dailyKey }
    if (-not $Force -and $exists) {
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Fetching files list..."
    try {
        $folders = @(Invoke-CanvasGet "/courses/$courseId/folders")
        $files = @(Invoke-CanvasGet "/courses/$courseId/files" @{ 'per_page' = '500' })
        
        if ($files.Count -eq 0) { return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; SkipReason = 'no files' } }

        $totalCourseSize = 0
        foreach ($f in $files) { $totalCourseSize += [long]$f.size }
        
        # Bin Packing limits. Adjusted to 10GB/20GB chunks for safer high concurrency.
        $SplitThreshold = 20GB
        $MaxChunkSize = 10GB
        $maxBlockSize = if ($totalCourseSize -gt $SplitThreshold) { $MaxChunkSize } else { $totalCourseSize + 1GB }

        $folderMap = @{}
        foreach ($f in $folders) { $folderMap["$($f.id)"] = $f.full_name }

        $currentBlock = @()
        $currentBlockSize = 0
        $blockNumber = 1
        $totalItemsProcessed = 0

        function Process-FileBlock {
            param($BlockFiles, $BNum)
            if ($BlockFiles.Count -eq 0) { return }

            $localDir = Join-Path $TempDir "files-$courseId-$([Guid]::NewGuid())"
            $zipPath = Join-Path $TempDir "files-$courseId-part$BNum-$([Guid]::NewGuid()).zip"

            try {
                # Safeguard to block jobs from proceeding if global cached disk space is > 100GB
                $maxCacheBytes = 100GB
                while ($true) {
                    $currentCacheBytes = 0
                    if (Test-Path $TempDir) {
                        $currentCacheBytes = (Get-ChildItem $TempDir -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    }
                    if ($null -eq $currentCacheBytes) { $currentCacheBytes = 0 }
                    if ($currentCacheBytes -lt $maxCacheBytes) { break }
                    Start-Sleep -Seconds 10
                }

                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                
                foreach ($bf in $BlockFiles) {
                    if (-not $bf.url) { continue }
                    $fId = "$($bf.folder_id)"
                    $path = if ($folderMap.ContainsKey($fId)) { $folderMap[$fId] } else { "" }
                    $path = $path -replace '(?i)^course files[/\\]?', ''

                    $targetDir = Join-Path $localDir $path
                    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                    
                    try {
                        Invoke-WebRequest -Uri $bf.url -OutFile (Join-Path $targetDir $bf.display_name) -UseBasicParsing -TimeoutSec 60
                    } catch {
                        Write-Warn "[$courseId] Skipped broken/missing file: $($bf.display_name)"
                    }
                }

                if ($totalCourseSize -gt $SplitThreshold) { Write-Info "[$courseId] Zipping part $BNum..." } else { Write-Info "[$courseId] Zipping files..." }
                
                # Using Native .NET ZipFile to bypass 2GB Compress-Archive bug
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::CreateFromDirectory($localDir, $zipPath, [System.IO.Compression.CompressionLevel]::Fastest, $false)

                $blockKey = if ($totalCourseSize -gt $SplitThreshold) { $dailyKey -replace '\.zip$', "-part$BNum.zip" } else { $dailyKey }
                $blockExt = if ($totalCourseSize -gt $SplitThreshold) { "-part$BNum.zip" } else { ".zip" }

                Publish-FileToS3 -LocalPath $zipPath -Key $blockKey
                Invoke-TieredPromotion -BaseKey $filesBaseKey -Today (Get-Date -Format 'yyyy-MM-dd') -Ext $blockExt
            } finally {
                # This guarantees cleanup runs NO MATTER WHAT happens above
                Remove-TempPath -Path $localDir -Recurse
                Remove-TempPath -Path $zipPath
            }
        }

        foreach ($file in $files) {
            $fSize = [long]$file.size
            if (($currentBlockSize + $fSize) -gt $maxBlockSize -and $currentBlock.Count -gt 0) {
                Process-FileBlock -BlockFiles $currentBlock -BNum $blockNumber
                $totalItemsProcessed += $currentBlock.Count
                $currentBlock = @()
                $currentBlockSize = 0
                $blockNumber++
            }
            $currentBlock += $file
            $currentBlockSize += $fSize
        }

        if ($currentBlock.Count -gt 0) {
            Process-FileBlock -BlockFiles $currentBlock -BNum $blockNumber
            $totalItemsProcessed += $currentBlock.Count
        }

        Write-Ok "[$courseId] Files ($totalItemsProcessed items) -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{ CourseId = $courseId; Action = 'written'; S3Key = $dailyKey; Items = $totalItemsProcessed }
    }
    catch {
        Write-Fail "[$courseId] Files fetch/zip failed: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }
}

function Format-CsvField {
    param([string]$Value)
    if ($null -eq $Value) { $Value = '' }
    '"' + ($Value -replace '"', '""') + '"'
}

function Backup-CourseGradebook {
    param($ExportData, [hashtable]$S3Cache = $null)
    Set-StrictMode -Off

    $courseId = $ExportData.CourseId
    $dailyKey = $ExportData.DailyGrades
    $gbBaseKey = "$($ExportData.BaseKey)-gradebook"

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Gradebook -> s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }

    $exists = $false
    if ($null -ne $S3Cache) { $exists = $S3Cache.ContainsKey($dailyKey) } else { $exists = Test-S3KeyExists -Key $dailyKey }
    if (-not $Force -and $exists) {
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Fetching gradebook data..."
    try {
        $assignments = @(Invoke-CanvasGet "/courses/$courseId/assignments")
        $enrollments = @(Invoke-CanvasGet "/courses/$courseId/enrollments" @{ 'type[]' = 'StudentEnrollment'; 'include[]' = 'user,grades' })
        $submissions = @(Invoke-CanvasGet "/courses/$courseId/students/submissions" @{ 'student_ids[]' = 'all'; 'per_page' = '100' })
        $asgnGroups = @(Invoke-CanvasGet "/courses/$courseId/assignment_groups" @{ 'include[]' = 'assignments' })
    }
    catch {
        Write-Fail "[$courseId] Gradebook fetch failed: $_"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }

    if ($enrollments.Count -eq 0) { return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; SkipReason = 'no students' } }

    $subLookup = @{}
    foreach ($sub in $submissions) {
        $key = ($sub.user_id.ToString()) + "_" + ($sub.assignment_id.ToString())
        $subLookup[$key] = if ($null -ne $sub.score) { "$($sub.score)" } else { '' }
    }

    $groupNames = @{}
    foreach ($grp in $asgnGroups) { $groupNames["$($grp.id)"] = $grp.name }

    $header = @(
        'Student', 'ID', 'SIS User ID', 'SIS Login ID', 'Section',
        'Current Score', 'Unposted Current Score', 'Final Score', 'Unposted Final Score',
        'Current Grade', 'Unposted Current Grade', 'Final Grade', 'Unposted Final Grade'
    )

    foreach ($asgn in $assignments) {
        $gName = if ($groupNames.ContainsKey("$($asgn.assignment_group_id)")) { $groupNames["$($asgn.assignment_group_id)"] } else { '' }
        $header += "$($asgn.name) ($gName) [$($asgn.id)]"
    }

    $lines = @()
    $lines += (($header | ForEach-Object { Format-CsvField $_ }) -join ',')

    foreach ($enr in $enrollments) {
        $user = $enr.user
        $grades = $enr.grades
        $sid = "$($enr.user_id)"

        $row = @(
            $(if ($user -and $user.sortable_name) { $user.sortable_name } else { '' }),
            $sid,
            $(if ($user -and $user.sis_user_id) { $user.sis_user_id } else { '' }),
            $(if ($user -and $user.login_id) { $user.login_id } else { '' }),
            $(if ($enr.sis_section_id) { $enr.sis_section_id } else { '' }),
            $(if ($grades -and $null -ne $grades.current_score) { "$($grades.current_score)" } else { '' }),
            $(if ($grades -and $null -ne $grades.unposted_current_score) { "$($grades.unposted_current_score)" } else { '' }),
            $(if ($grades -and $null -ne $grades.final_score) { "$($grades.final_score)" } else { '' }),
            $(if ($grades -and $null -ne $grades.unposted_final_score) { "$($grades.unposted_final_score)" } else { '' }),
            $(if ($grades -and $grades.current_grade) { $grades.current_grade } else { '' }),
            $(if ($grades -and $grades.unposted_current_grade) { $grades.unposted_current_grade } else { '' }),
            $(if ($grades -and $grades.final_grade) { $grades.final_grade } else { '' }),
            $(if ($grades -and $grades.unposted_final_grade) { $grades.unposted_final_grade } else { '' })
        )

        foreach ($asgn in $assignments) {
            $subKey = $sid + "_" + $asgn.id
            $score = if ($subLookup.ContainsKey($subKey)) { $subLookup[$subKey] } else { '' }
            $row += $score
        }
        $lines += (($row | ForEach-Object { Format-CsvField $_ }) -join ',')
    }

    $bom = [System.Text.Encoding]::UTF8.GetPreamble()
    $csvText = $lines -join "`r`n"
    $csvBytes = [System.Text.Encoding]::UTF8.GetBytes($csvText)
    $allBytes = $bom + $csvBytes
    $tmpCsv = Join-Path $TempDir "canvas-gradebook-$courseId-$([Guid]::NewGuid()).csv"

    try {
        [System.IO.File]::WriteAllBytes($tmpCsv, $allBytes)
        Publish-FileToS3 -LocalPath $tmpCsv -Key $dailyKey
        Invoke-TieredPromotion -BaseKey $gbBaseKey -Today (Get-Date -Format 'yyyy-MM-dd') -Ext '.csv'

        Write-Ok "[$courseId] Gradebook ($($enrollments.Count) students) -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{ CourseId = $courseId; Action = 'written'; S3Key = $dailyKey; Students = $enrollments.Count }
    }
    catch {
        Write-Fail "[$courseId] Gradebook upload failed: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }
    finally {
        Remove-TempPath -Path $tmpCsv
    }
}

# ============================================================
# Persistent sub-account backup
# ============================================================
function Backup-SubAccount {
    param([string]$AccountId)
    Set-StrictMode -Off

    Write-Info "=== Sub-account backup: account $AccountId ==="
    $accountInfo = (Invoke-CanvasGet "/accounts/${AccountId}" | Select-Object -First 1)
    $accountSlug = ConvertTo-CanvasSlug (Get-JsonValue $accountInfo 'name')
    if (-not $accountSlug) { $accountSlug = "account-$AccountId" }
    $basePrefix = "canvas-backups/persistent/$accountSlug"

    $courses = @(Invoke-CanvasGet "/accounts/${AccountId}/courses" @{ 'include[]' = 'term,account_name'; 'per_page' = '100' })
    $courses = @($courses | Group-Object id | ForEach-Object { $_.Group[0] })

    Write-Info "Sub-account '$($accountInfo.name)': $($courses.Count) courses"

    $contentResults = @()
    foreach ($course in $courses) {
        $cTerm = Get-JsonValue $course 'term'
        $termName = if ($cTerm) { Get-JsonValue $cTerm 'name' } else { '' }
        $startAt = if ($cTerm) { Get-JsonValue $cTerm 'start_at' } else { '' }
        $termParts = Get-TermParts -TermName $termName -StartAt $startAt
        
        $acctSlug = ConvertTo-CanvasSlug (Get-JsonValue $course 'account_name')
        $courseCode = Get-JsonValue $course 'course_code'
        $deptSlug = Get-DeptSlug -AccountSlug $acctSlug -CourseCode $courseCode
        
        $baseKey = Get-CourseS3BaseKey -Course $course -TermParts $termParts -DeptSlug $deptSlug -BasePrefix $basePrefix
        $todayString = Get-Date -Format 'yyyy-MM-dd'

        $exportData = [pscustomobject]@{
            CourseId     = $course.id
            BaseKey      = $baseKey
            DailyContent = "$baseKey/daily/$todayString.imscc"
            DailyPages   = "$baseKey-pages/daily/$todayString.json"
            DailyFiles   = "$baseKey-files/daily/$todayString.zip"
            DailyGrades  = "$baseKey-gradebook/daily/$todayString.csv"
        }

        if (-not $SkipContent) { $contentResults += Backup-CourseContent -ExportData $exportData }
        if (-not $SkipGradebooks) { Backup-CourseGradebook -ExportData $exportData | Out-Null }
        if (-not $SkipPages) { Backup-CoursePages -ExportData $exportData | Out-Null }
        if (-not $SkipFiles) { Backup-CourseFiles -ExportData $exportData | Out-Null }
    }

    $written = @($contentResults | Where-Object { $_.Action -eq 'written' }).Count
    Write-Ok "Sub-account '$($accountInfo.name)' Content Exports: $written written."
}

# ============================================================
# Coverage verification
# ============================================================
function Test-BackupCoverage {
    param([string[]]$TermIds, [string]$Year, [string]$Semester)

    Write-Info "=== Verifying coverage for $Year/$Semester ==="
    $canvasCourses = @()
    foreach ($tid in $TermIds) {
        $canvasCourses += @(Invoke-CanvasGet "/accounts/${RootAccountId}/courses" @{ enrollment_term_id = $tid })
    }

    $canvasCourses = @($canvasCourses | Group-Object id | ForEach-Object { $_.Group[0] })
    $canvasCount = $canvasCourses.Count

    $prefix = "canvas-backups/$Year/$Semester/"
    $keys = @(Get-S3KeysUnderPrefix -Prefix $prefix)
    $s3Count = 0

    if ($keys.Count -gt 0) {
        $unique = $keys | ForEach-Object {
            $parts = $_ -split '/'
            if ($parts.Length -ge 5) { ($parts[0..4] -join '/') }
        } | Where-Object { $_ } | Sort-Object -Unique

        $s3Count = @($unique).Count
    }

    $coveragePct = if ($canvasCount -gt 0) { [int]([Math]::Round(($s3Count / $canvasCount) * 100)) } else { 0 }

    if ($coveragePct -lt 90) { Write-Fail "BACKUP FAILURE: coverage_percent=$coveragePct  canvas_courses=$canvasCount  s3_courses=$s3Count" }
    else { Write-Ok "BACKUP OK:      coverage_percent=$coveragePct  canvas_courses=$canvasCount  s3_courses=$s3Count" }

    [pscustomobject]@{ CoveragePercent = $coveragePct; CanvasCourses = $canvasCount; S3Courses = $s3Count }
}

# ============================================================
# Main orchestration
# ============================================================
function Main {
    $runStamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $todayString = Get-Date -Format 'yyyy-MM-dd'
    $yearMonth = Get-Date -Format 'yyyy/MM'
    $logFileName = "backup-$runStamp.log"
    $summaryFileName = "backup-$runStamp.summary.txt"
    $transcriptPath = Join-Path $LogDir $logFileName
    $s3LogKey = "canvas-backups/logs/$yearMonth/$logFileName"
    $s3SummaryKey = "canvas-backups/logs/$yearMonth/$summaryFileName"

    $transcriptStarted = $false
    $cW = 0; $cS = 0; $cF = 0; $cD = 0
    $pW = 0; $fW = 0; $gW = 0;
    $coverage = $null
    $runStatus = 'FAILED'
    $startTime = Get-Date

    $today = Get-Date
    $thisYear = $today.Year
    $prevYear = $today.Year - 1
    $accountCache = @{}
    $termMap = @{}
    $allCourses = @()
    $activeTerms = @()
    $termIds = @()

    if (-not (Get-Command aws -ErrorAction SilentlyContinue)) { throw "AWS CLI (aws.exe) not found in PATH." }
    if (-not $CanvasBaseUrl) { throw "CANVAS_BASE_URL is required" }
    if (-not $CanvasApiToken) { throw "CANVAS_API_TOKEN is required" }
    if (-not $S3Bucket) { throw "S3_BUCKET is required" }

    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    if (-not (Test-Path $TempDir)) { 
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null 
    } else {
        # Auto-clean orphaned files from previous failed runs to prevent immediate deadlocks
        Remove-Item -Path "$TempDir\*" -Recurse -Force -ErrorAction SilentlyContinue
    }

    try {
        Start-Transcript -Path $transcriptPath | Out-Null
        $transcriptStarted = $true

        Write-Info "=== Canvas Nightly Backup starting at $startTime ==="
        Write-Info "Canvas: $CanvasBaseUrl"
        Write-Info "S3: s3://$S3Bucket (region: $AwsRegion)"
        Write-Info "Root Account ID: $RootAccountId"
        Write-Info "Local TempDir: $TempDir"
        Write-Info "Concurrent Jobs: $ConcurrentJobs"

        if ($Force) { Write-Warn "Force mode - same-day dedup bypassed" }
        if ($CourseId) { Write-Warn "Test mode: filtering to CourseId=$CourseId" }
        if ($MaxCourses -gt 0) { Write-Warn "Test mode: limiting to first $MaxCourses course(s)" }
        if ($WhatIfMode) { Write-Warn "DRY RUN mode enabled" }

        # ------------------------------------------------------------
        # Gather Active Term Courses
        # ------------------------------------------------------------
        if ($CourseId) {
            $singleCourse = Get-CanvasCourseById -CourseId $CourseId
            if (-not $singleCourse) { Write-Warn "CourseId $CourseId not found"; return }
            $allCourses = @($singleCourse)
            Write-Info "Single-course mode enabled: CourseId=$CourseId  matched=1"
        }
        else {
            Write-Info "=== Resolving enrollment terms ==="
            $allTerms = @(Get-EnrollmentTerms -AccountId $RootAccountId)
            $activeTerms = @( $allTerms | Where-Object { $_.start_at -and ([datetime]$_.start_at).Year -in @($thisYear, $prevYear) } )
            $termIds = @($activeTerms | ForEach-Object { "$($_.id)" })
            
            if ($activeTerms.Count -eq 0) { Write-Warn "No active terms found"; return }
            Write-Info "Active terms ($($activeTerms.Count)): $($termIds -join ', ')"

            foreach ($t in $activeTerms) { $termMap["$($t.id)"] = $t }

            Write-Info "=== Listing courses ==="
            foreach ($term in $activeTerms) {
                $tc = @(Invoke-CanvasGet "/accounts/${RootAccountId}/courses" @{ enrollment_term_id = $term.id; 'include[]' = 'term,account_name' })
                Write-Info "  '$($term.name)': $($tc.Count) courses"
                $allCourses += $tc
            }

            $allCourses = @($allCourses | Group-Object id | ForEach-Object { $_.Group[0] })
            Write-Info "Total unique courses before test filters: $($allCourses.Count)"

            if ($MaxCourses -gt 0) {
                $allCourses = @($allCourses | Select-Object -First $MaxCourses)
                Write-Info "MaxCourses mode enabled: processing first $($allCourses.Count) courses"
            }
            if ($allCourses.Count -eq 0) { Write-Warn "No courses matched filter"; return }
            Write-Info "Total unique courses after test filters: $($allCourses.Count)"
        }

        # ------------------------------------------------------------
        # Pre-Calculate S3 Paths (High Speed / Crash Proof)
        # ------------------------------------------------------------
        Write-Info "=== Pre-calculating S3 paths for all courses ==="
        $exportList = @()
        foreach ($c in $allCourses) {
            $termObj = Get-JsonValue $termMap "$($c.enrollment_term_id)"
            $cTerm = Get-JsonValue $c 'term'
            $termName = if ($termObj) { $termObj.name } elseif ($cTerm) { Get-JsonValue $cTerm 'name' } else { '' }
            $startAt  = if ($termObj) { $termObj.start_at } elseif ($cTerm) { Get-JsonValue $cTerm 'start_at' } else { '' }
            $tp = Get-TermParts -TermName $termName -StartAt $startAt

            $acctName = Get-JsonValue $c 'account_name'
            if ($acctName) {
                $acctSlug = ConvertTo-CanvasSlug $acctName
            } else {
                $acctKey = Get-JsonValue $c 'account_id'
                if ($acctKey -and -not $accountCache.ContainsKey("$acctKey")) {
                    try {
                        $acct = (Invoke-CanvasGet "/accounts/${acctKey}" | Select-Object -First 1)
                        $accountCache["$acctKey"] = ConvertTo-CanvasSlug (Get-JsonValue $acct 'name')
                    } catch { $accountCache["$acctKey"] = 'other' }
                }
                $acctSlug = if ($acctKey) { $accountCache["$acctKey"] } else { 'other' }
            }
            
            $courseCode = Get-JsonValue $c 'course_code'
            $dept = Get-DeptSlug -AccountSlug $acctSlug -CourseCode $courseCode
            $baseKey = Get-CourseS3BaseKey -Course $c -TermParts $tp -DeptSlug $dept

            $exportList += [pscustomobject]@{
                CanvasCourse = $c
                CourseId     = $c.id
                BaseKey      = $baseKey
                DailyContent = "$baseKey/daily/$todayString.imscc"
                DailyPages   = "$baseKey-pages/daily/$todayString.json"
                DailyFiles   = "$baseKey-files/daily/$todayString.zip"
                DailyGrades  = "$baseKey-gradebook/daily/$todayString.csv"
            }
        }

        # ------------------------------------------------------------
        # Build High-Speed S3 Deduplication Cache
        # ------------------------------------------------------------
        $s3Cache = @{}
        if (-not $Force -and -not $WhatIfMode) {
            Write-Info "=== Building S3 deduplication cache for active terms ==="
            $activePrefixes = @($exportList | Select-Object -ExpandProperty BaseKey | ForEach-Object { 
                $parts = $_ -split '/'
                $parts[0..2] -join '/' + '/' 
            } | Sort-Object -Unique)

            foreach ($p in $activePrefixes) {
                Write-Info "Fetching existing S3 keys under: $p"
                $keys = Get-S3KeysUnderPrefix -Prefix $p
                foreach ($k in $keys) {
                    if ($k -match "daily/$todayString") { $s3Cache[$k] = $true }
                }
            }
            Write-Info "Cached $($s3Cache.Count) existing daily backups for today."
        }

        # ------------------------------------------------------------
        # Content export pass (.imscc Background Jobs)
        # ------------------------------------------------------------
        $contentResults = @()
        if ($SkipContent) {
            Write-Warn "Skipping course content export (.imscc) pass"
        }
        else {
            Write-Info "=== Backing up course content exports ($ConcurrentJobs BACKGROUND JOBS) ==="
            
            $exportJobBlock = {
                param($ArgsHash)
                Set-StrictMode -Off
                
                $CanvasApiToken = $ArgsHash.CanvasApiToken
                $CanvasBaseUrl = $ArgsHash.CanvasBaseUrl
                $S3Bucket = $ArgsHash.S3Bucket
                $AwsRegion = $ArgsHash.AwsRegion
                $TempDir = $ArgsHash.TempDir
                $PollTimeoutMins = $ArgsHash.PollTimeoutMins
                $PollIntervalSecs = $ArgsHash.PollIntervalSecs
                $todayString = $ArgsHash.TodayString

                function Remove-TempPath {
                    param([string]$Path, [switch]$Recurse)
                    if (-not (Test-Path $Path)) { return }
                    Start-Sleep -Seconds 1 
                    $delRetries = 90
                    while ((Test-Path $Path) -and $delRetries -gt 0) {
                        if ($Recurse) { Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue }
                        else { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
                        if (Test-Path $Path) { Start-Sleep -Seconds 2; $delRetries-- }
                    }
                }

                function Get-JsonValue {
                    param($Obj, [string]$Property)
                    if ($null -eq $Obj) { return $null }
                    if ($Obj -is [hashtable]) { return $Obj[$Property] }
                    if ($Obj.PSObject.Properties.Match($Property).Count -gt 0) { return $Obj.$Property }
                    return $null
                }

                function Invoke-CanvasApiPage {
                    param([string]$Url)
                    $attempt = 0
                    while ($true) {
                        $attempt++
                        try { return Invoke-WebRequest -Uri $Url -Method Get -Headers @{ Authorization = "Bearer $CanvasApiToken" } -UseBasicParsing -TimeoutSec 60 } 
                        catch [System.Net.WebException] {
                            $response = $_.Exception.Response
                            if ($response -and $response.StatusCode -eq 429) {
                                $ra = 10
                                if ($response.Headers['Retry-After']) { try { $ra = [int]$response.Headers['Retry-After'] } catch {} }
                                Start-Sleep -Seconds $ra
                            } elseif ($attempt -lt 3) { Start-Sleep -Seconds (5 * $attempt) } else { throw }
                        }
                    }
                }

                function Invoke-CanvasGet {
                    param([string]$Endpoint)
                    $base = $CanvasBaseUrl.TrimEnd('/')
                    $url = "$base/api/v1$Endpoint"
                    if ($url -notmatch 'per_page') { $url += if ($url -match '\?') { '&per_page=100' } else { '?per_page=100' } }
                    $results = @()
                    do {
                        $resp = Invoke-CanvasApiPage -Url $url
                        $page = $resp.Content | ConvertFrom-Json
                        if ($null -ne $page) { $results += @($page) }
                        $url = $null
                        $linkHeader = $resp.Headers['Link']
                        if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) { $url = $Matches[1] }
                    } while ($url)
                    return $results
                }

                function Invoke-CanvasPost {
                    param([string]$Endpoint, [hashtable]$Body)
                    $uri = "$($CanvasBaseUrl.TrimEnd('/'))/api/v1$Endpoint"
                    $json = $Body | ConvertTo-Json -Compress
                    $resp = Invoke-WebRequest -Uri $uri -Method Post -Headers @{ Authorization = "Bearer $CanvasApiToken" } -ContentType 'application/json' -Body $json -UseBasicParsing -TimeoutSec 60
                    return $resp.Content | ConvertFrom-Json
                }

                function Publish-FileToS3 {
                    param([string]$LocalPath, [string]$Key, [string]$StorageClass = 'GLACIER_IR')
                    $maxRetries = 3
                    $attempt = 0
                    $success = $false
                    while ($attempt -lt $maxRetries -and -not $success) {
                        $attempt++
                        & aws s3 cp $LocalPath "s3://$S3Bucket/$Key" --storage-class $StorageClass --region $AwsRegion --no-progress
                        if ($LASTEXITCODE -eq 0) { $success = $true } else { if ($attempt -lt $maxRetries) { $sleepTime = $attempt * 10; Start-Sleep -Seconds $sleepTime } }
                    }
                    if (-not $success) { throw "aws s3 cp failed" }
                }

                function Invoke-S3CopyObject {
                    param([string]$SourceKey, [string]$DestKey)
                    & aws s3 cp "s3://$S3Bucket/$SourceKey" "s3://$S3Bucket/$DestKey" --metadata-directive COPY --storage-class GLACIER_IR --region $AwsRegion --no-progress
                    if ($LASTEXITCODE -ne 0) { throw "aws s3 copy failed" }
                }

                function Get-S3ObjectsUnderPrefix {
                    param([string]$Prefix)
                    $results = @()
                    $token = $null
                    $oldEAP = $ErrorActionPreference
                    $ErrorActionPreference = 'Continue'
                    do {
                        $awsArgs = @('s3api', 'list-objects-v2', '--bucket', $S3Bucket, '--prefix', $Prefix, '--output', 'json', '--region', $AwsRegion)
                        if ($token) { $awsArgs += @('--continuation-token', $token) }
                        $json = & aws @awsArgs 2>$null
                        if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }
                        $page = $json | ConvertFrom-Json
                        if ($null -eq $page) { break }
                        
                        if (Get-JsonValue $page 'Contents') {
                            $results += @($page.Contents | Select-Object Key, LastModified)
                        }
                        $isTrunc = Get-JsonValue $page 'IsTruncated'
                        $nxtTok = Get-JsonValue $page 'NextContinuationToken'
                        if ($isTrunc -and $nxtTok) { $token = $page.NextContinuationToken } else { $token = $null }
                    } while ($token)
                    $ErrorActionPreference = $oldEAP
                    if (-not $results) { return @() }
                    return @($results | Sort-Object { [datetime]$_.LastModified } -Descending)
                }

                function Invoke-S3Prune {
                    param([string]$Prefix, [int]$KeepCount, [string]$Ext)
                    $objects = @(Get-S3ObjectsUnderPrefix -Prefix $Prefix | Where-Object { $_.Key.EndsWith($Ext) })
                    if ($objects.Count -le $KeepCount) { return }
                    $toDelete = @($objects | Select-Object -Skip $KeepCount)
                    foreach ($obj in $toDelete) {
                        & aws s3api delete-object --bucket $S3Bucket --key $obj.Key --region $AwsRegion | Out-Null
                    }
                }

                function Invoke-TieredPromotion {
                    param([string]$BaseKey, [string]$Today)
                    $dailyKey = "$BaseKey/daily/$Today.imscc"
                    $weeklyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/weekly/")
                    $monthlyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/monthly/")

                    $daysSinceWeekly = 999
                    $daysSinceMonthly = 999

                    if ($weeklyObjs.Count -gt 0) { $daysSinceWeekly = ((Get-Date) - [datetime]$weeklyObjs[0].LastModified).TotalDays }
                    if ($monthlyObjs.Count -gt 0) { $daysSinceMonthly = ((Get-Date) - [datetime]$monthlyObjs[0].LastModified).TotalDays }

                    if ($daysSinceWeekly -ge 7) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/weekly/$Today.imscc" }
                    if ($daysSinceWeekly -ge 7 -and $daysSinceMonthly -ge 30) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/monthly/$Today.imscc" }

                    Invoke-S3Prune -Prefix "$BaseKey/daily/" -KeepCount 14 -Ext '.imscc'
                    Invoke-S3Prune -Prefix "$BaseKey/weekly/" -KeepCount 5 -Ext '.imscc'
                    Invoke-S3Prune -Prefix "$BaseKey/monthly/" -KeepCount 6 -Ext '.imscc'
                }

                $activeJobs = @()
                foreach ($c in $ArgsHash.Chunk) {
                    Write-Output [pscustomobject]@{ CourseId = $c.CourseId; Action = 'submitted' }
                    try {
                        $job = Invoke-CanvasPost "/courses/$($c.CourseId)/content_exports" @{ export_type = 'common_cartridge'; skip_notifications = $true }
                        $activeJobs += [pscustomobject]@{ CourseId = $c.CourseId; JobId = $job.id; DailyKey = $c.DailyContent; BaseKey = $c.BaseKey; SubmittedAt = (Get-Date) }
                    } catch {
                        Write-Output [pscustomobject]@{ CourseId = $c.CourseId; Action = 'failed'; Error = $_.Exception.Message }
                    }
                }

                while ($activeJobs.Count -gt 0) {
                    Start-Sleep -Seconds $PollIntervalSecs
                    $stillActive = @()

                    foreach ($job in $activeJobs) {
                        $courseId = $job.CourseId

                        if (((Get-Date) - $job.SubmittedAt).TotalMinutes -gt $PollTimeoutMins) {
                            Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Poll timeout' }
                            continue
                        }

                        try { $exp = (Invoke-CanvasGet "/courses/$courseId/content_exports/$($job.JobId)" | Select-Object -First 1) } catch { $stillActive += $job; continue }
                        $wfState = Get-JsonValue $exp 'workflow_state'
                        
                        if ($wfState -eq 'failed') { Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Canvas export failed' }; continue }

                        if ($wfState -eq 'exported') {
                            $attachment = Get-JsonValue $exp 'attachment'
                            $downloadUrl = Get-JsonValue $attachment 'url'
                            if (-not $downloadUrl) { Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Missing URL' }; continue }

                            $tempFile = Join-Path $TempDir "canvas-export-$courseId-$([Guid]::NewGuid()).imscc"
                            try {
                                Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'downloading' }
                                Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 600
                                $sizeBytes = (Get-Item $tempFile).Length
                                
                                Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'uploading' }
                                Publish-FileToS3 -LocalPath $tempFile -Key $job.DailyKey
                                Invoke-TieredPromotion -BaseKey $job.BaseKey -Today $todayString
                                
                                Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'written'; SizeBytes = $sizeBytes }
                            } catch { Write-Output [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = $_.Exception.Message }
                            } finally { 
                                Remove-TempPath -Path $tempFile 
                            }
                        } else {
                            $stillActive += $job
                        }
                    }
                    $activeJobs = $stillActive
                }
            }

            # Filter Export List 
            $coursesToExport = @()
            foreach ($c in $exportList) {
                if ($WhatIfMode) {
                    Write-DryRun "[$($c.CourseId)] Would submit Canvas content export"
                    $contentResults += [pscustomobject]@{ CourseId = $c.CourseId; Action = 'dryrun' }
                    continue
                }
                if (-not $Force -and $s3Cache.ContainsKey($c.DailyContent)) {
                    $contentResults += [pscustomobject]@{ CourseId = $c.CourseId; Action = 'skipped' }
                    continue
                }
                $coursesToExport += $c
            }

            # Group into chunks of 5 for Background Jobs
            $chunks = @()
            for ($i = 0; $i -lt $coursesToExport.Count; $i += 5) {
                $chunks += ,@($coursesToExport[$i..[math]::Min($i+4, $coursesToExport.Count - 1)])
            }

            # Orchestrate Background Jobs (Maintain $ConcurrentJobs active threads)
            $chunkIndex = 0
            while ($true) {
                $jobs = Get-Job -Command "*CanvasExport*" 2>$null
                
                foreach ($j in $jobs) {
                    $results = Receive-Job -Job $j
                    if ($null -eq $results) { continue }
                    foreach ($res in $results) {
                        if ($res.Action -eq 'submitted') { Write-Info "[$($res.CourseId)] Submitting export job to Canvas..." }
                        elseif ($res.Action -eq 'downloading') { Write-Info "[$($res.CourseId)] Ready! Downloading..." }
                        elseif ($res.Action -eq 'uploading') { Write-Info "[$($res.CourseId)] Uploading to S3..." }
                        elseif ($res.Action -eq 'written') {
                            Write-Ok "[$($res.CourseId)] $([Math]::Round($res.SizeBytes / 1MB, 1)) MB -> S3"
                            $contentResults += $res
                        }
                        elseif ($res.Action -eq 'failed') {
                            Write-Fail "[$($res.CourseId)] Export failed: $($res.Error)"
                            $contentResults += $res
                        }
                    }
                }

                $runningJobs = @($jobs | Where-Object State -eq 'Running')
                if ($chunkIndex -ge $chunks.Count -and $runningJobs.Count -eq 0) { break }

                while ($chunkIndex -lt $chunks.Count -and $runningJobs.Count -lt $ConcurrentJobs) {
                    $chunk = $chunks[$chunkIndex]
                    $jobArgs = @{
                        Chunk = $chunk
                        CanvasApiToken = $CanvasApiToken
                        CanvasBaseUrl = $CanvasBaseUrl
                        S3Bucket = $S3Bucket
                        AwsRegion = $AwsRegion
                        TempDir = $TempDir
                        PollTimeoutMins = $PollTimeoutMins
                        PollIntervalSecs = $PollIntervalSecs
                        TodayString = $todayString
                    }
                    Start-Job -Name "CanvasExport-$chunkIndex" -ScriptBlock $exportJobBlock -ArgumentList $jobArgs | Out-Null
                    $chunkIndex++
                    $jobs = Get-Job -Command "*CanvasExport*" 2>$null
                    $runningJobs = @($jobs | Where-Object State -eq 'Running')
                }

                $jobs | Where-Object State -ne 'Running' | Remove-Job
                Start-Sleep -Seconds 2
            }
        }
        $cW = @($contentResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count

        # ------------------------------------------------------------
        # Generic Worker Job Block (Pages, Files, Gradebooks)
        # ------------------------------------------------------------
        $DataWorkerJobBlock = {
            param($ArgsHash)
            Set-StrictMode -Off

            $TaskType = $ArgsHash.TaskType
            $CanvasApiToken = $ArgsHash.CanvasApiToken
            $CanvasBaseUrl = $ArgsHash.CanvasBaseUrl
            $S3Bucket = $ArgsHash.S3Bucket
            $AwsRegion = $ArgsHash.AwsRegion
            $TempDir = $ArgsHash.TempDir
            $Force = $ArgsHash.Force
            $WhatIfMode = $ArgsHash.WhatIfMode
            $TodayString = $ArgsHash.TodayString

            # --- INJECTED HELPER FUNCTIONS ---
            function Write-Info {} function Write-Warn {} function Write-Ok {} function Write-Fail {} function Write-DryRun {}
            function Remove-TempPath { param([string]$Path, [switch]$Recurse) if (-not (Test-Path $Path)) { return }; Start-Sleep -Seconds 1; $delRetries = 90; while ((Test-Path $Path) -and $delRetries -gt 0) { if ($Recurse) { Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue } else { Remove-Item $Path -Force -ErrorAction SilentlyContinue }; if (Test-Path $Path) { Start-Sleep -Seconds 2; $delRetries-- } } }
            function Get-JsonValue { param($Obj, [string]$Property) if ($null -eq $Obj) { return $null }; if ($Obj -is [hashtable]) { return $Obj[$Property] }; if ($Obj.PSObject.Properties.Match($Property).Count -gt 0) { return $Obj.$Property }; return $null }
            function Invoke-CanvasApiPage { param([string]$Url) $attempt = 0; while ($true) { $attempt++; try { return Invoke-WebRequest -Uri $Url -Method Get -Headers @{ Authorization = "Bearer $CanvasApiToken" } -UseBasicParsing -TimeoutSec 60 } catch [System.Net.WebException] { $response = $_.Exception.Response; if ($response -and $response.StatusCode -eq 429) { $ra = 10; if ($response.Headers['Retry-After']) { try { $ra = [int]$response.Headers['Retry-After'] } catch {} }; Start-Sleep -Seconds $ra } elseif ($attempt -lt 3) { Start-Sleep -Seconds (5 * $attempt) } else { throw } } } }
            function Invoke-CanvasGet { param([string]$Endpoint, [hashtable]$Query = @{}) $base = $CanvasBaseUrl.TrimEnd('/'); $qs = @(); $hasPerPage = $false; foreach ($kv in $Query.GetEnumerator()) { if ($kv.Key -eq 'per_page') { $hasPerPage = $true }; $qs += "$([Uri]::EscapeDataString($kv.Key))=$([Uri]::EscapeDataString([string]$kv.Value))" }; if (-not $hasPerPage) { $qs += 'per_page=100' }; $url = "$base/api/v1${Endpoint}?" + ($qs -join '&'); $results = @(); do { $resp = Invoke-CanvasApiPage -Url $url; $page = $resp.Content | ConvertFrom-Json; if ($null -ne $page) { $results += @($page) }; $url = $null; $linkHeader = $resp.Headers['Link']
                    if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) { $url = $Matches[1] } } while ($url); return $results }
            
            function Publish-FileToS3 {
                param([string]$LocalPath, [string]$Key, [string]$StorageClass = 'GLACIER_IR')
                if ($WhatIfMode) { return }
                $maxRetries = 3; $attempt = 0; $success = $false
                while ($attempt -lt $maxRetries -and -not $success) {
                    $attempt++
                    & aws s3 cp $LocalPath "s3://$S3Bucket/$Key" --storage-class $StorageClass --region $AwsRegion --no-progress
                    if ($LASTEXITCODE -eq 0) { $success = $true } else { if ($attempt -lt $maxRetries) { $sleepTime = $attempt * 10; Start-Sleep -Seconds $sleepTime } }
                }
                if (-not $success) { throw "aws s3 cp failed" }
            }
            
            function Publish-TextToS3 { param([string]$Text, [string]$Key) $tmpFile = Join-Path $TempDir "s3-text-$([Guid]::NewGuid()).txt"; try { [System.IO.File]::WriteAllText($tmpFile, $Text, [System.Text.Encoding]::UTF8); Publish-FileToS3 -LocalPath $tmpFile -Key $Key -StorageClass 'STANDARD' } finally { Remove-TempPath -Path $tmpFile } }
            function Invoke-S3CopyObject { param([string]$SourceKey, [string]$DestKey) if ($WhatIfMode) { return }; & aws s3 cp "s3://$S3Bucket/$SourceKey" "s3://$S3Bucket/$DestKey" --metadata-directive COPY --storage-class GLACIER_IR --region $AwsRegion --no-progress; if ($LASTEXITCODE -ne 0) { throw "aws s3 copy failed" } }
            function Get-S3ObjectsUnderPrefix { param([string]$Prefix) $results = @(); $token = $null; $oldEAP = $ErrorActionPreference; $ErrorActionPreference = 'Continue'; do { $awsArgs = @('s3api', 'list-objects-v2', '--bucket', $S3Bucket, '--prefix', $Prefix, '--output', 'json', '--region', $AwsRegion); if ($token) { $awsArgs += @('--continuation-token', $token) }; $json = & aws @awsArgs 2>$null; if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }; $page = $json | ConvertFrom-Json; if ($null -eq $page) { break }; if (Get-JsonValue $page 'Contents') { $results += @($page.Contents | Select-Object Key, LastModified) }; $isTrunc = Get-JsonValue $page 'IsTruncated'; $nxtTok = Get-JsonValue $page 'NextContinuationToken'; if ($isTrunc -and $nxtTok) { $token = $page.NextContinuationToken } else { $token = $null } } while ($token); $ErrorActionPreference = $oldEAP; if (-not $results) { return @() }; return @($results | Sort-Object { [datetime]$_.LastModified } -Descending) }
            function Invoke-S3Prune { param([string]$Prefix, [int]$KeepCount, [string]$Ext) $objects = @(Get-S3ObjectsUnderPrefix -Prefix $Prefix | Where-Object { $_.Key.EndsWith($Ext) }); if ($objects.Count -le $KeepCount) { return }; $toDelete = @($objects | Select-Object -Skip $KeepCount); foreach ($obj in $toDelete) { if (-not $WhatIfMode) { & aws s3api delete-object --bucket $S3Bucket --key $obj.Key --region $AwsRegion | Out-Null } } }
            function Invoke-TieredPromotion { param([string]$BaseKey, [string]$Ext) $dailyKey = "$BaseKey/daily/$TodayString$Ext"; $weeklyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/weekly/"); $monthlyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/monthly/"); $daysSinceWeekly = 999; $daysSinceMonthly = 999; if ($weeklyObjs.Count -gt 0) { $daysSinceWeekly = ((Get-Date) - [datetime]$weeklyObjs[0].LastModified).TotalDays }; if ($monthlyObjs.Count -gt 0) { $daysSinceMonthly = ((Get-Date) - [datetime]$monthlyObjs[0].LastModified).TotalDays }; if ($daysSinceWeekly -ge 7) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/weekly/$TodayString$Ext" }; if ($daysSinceWeekly -ge 7 -and $daysSinceMonthly -ge 30) { Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/monthly/$TodayString$Ext" }; Invoke-S3Prune -Prefix "$BaseKey/daily/" -KeepCount 14 -Ext $Ext; Invoke-S3Prune -Prefix "$BaseKey/weekly/" -KeepCount 5 -Ext $Ext; Invoke-S3Prune -Prefix "$BaseKey/monthly/" -KeepCount 6 -Ext $Ext }
            function Format-CsvField { param([string]$Value) if ($null -eq $Value) { $Value = '' }; '"' + ($Value -replace '"', '""') + '"' }

            # --- WORKER LOOP ---
            $output = @()
            foreach ($c in $ArgsHash.Chunk) {
                $courseId = $c.CourseId

                if ($TaskType -eq 'Gradebook') {
                    if ($WhatIfMode) { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun' }; continue }
                    try {
                        $assignments = @(Invoke-CanvasGet "/courses/$courseId/assignments")
                        $enrollments = @(Invoke-CanvasGet "/courses/$courseId/enrollments" @{ 'type[]' = 'StudentEnrollment'; 'include[]' = 'user,grades' })
                        if ($enrollments.Count -gt 0) {
                            $submissions = @(Invoke-CanvasGet "/courses/$courseId/students/submissions" @{ 'student_ids[]' = 'all'; 'per_page' = '100' })
                            $asgnGroups = @(Invoke-CanvasGet "/courses/$courseId/assignment_groups" @{ 'include[]' = 'assignments' })
                            $subLookup = @{}; foreach ($sub in $submissions) { $subLookup["$($sub.user_id)_$($sub.assignment_id)"] = if ($null -ne $sub.score) { "$($sub.score)" } else { '' } }
                            $groupNames = @{}; foreach ($grp in $asgnGroups) { $groupNames["$($grp.id)"] = $grp.name }
                            $header = @('Student', 'ID', 'SIS User ID', 'SIS Login ID', 'Section', 'Current Score', 'Unposted Current Score', 'Final Score', 'Unposted Final Score', 'Current Grade', 'Unposted Current Grade', 'Final Grade', 'Unposted Final Grade')
                            foreach ($asgn in $assignments) { $gName = if ($groupNames.ContainsKey("$($asgn.assignment_group_id)")) { $groupNames["$($asgn.assignment_group_id)"] } else { '' }; $header += "$($asgn.name) ($gName) [$($asgn.id)]" }
                            $lines = @((($header | ForEach-Object { Format-CsvField $_ }) -join ','))
                            foreach ($enr in $enrollments) {
                                $user = $enr.user; $grades = $enr.grades; $sid = "$($enr.user_id)"
                                $row = @( $(if ($user.sortable_name) { $user.sortable_name } else { '' }), $sid, $(if ($user.sis_user_id) { $user.sis_user_id } else { '' }), $(if ($user.login_id) { $user.login_id } else { '' }), $(if ($enr.sis_section_id) { $enr.sis_section_id } else { '' }), $(if ($null -ne $grades.current_score) { "$($grades.current_score)" } else { '' }), $(if ($null -ne $grades.unposted_current_score) { "$($grades.unposted_current_score)" } else { '' }), $(if ($null -ne $grades.final_score) { "$($grades.final_score)" } else { '' }), $(if ($null -ne $grades.unposted_final_score) { "$($grades.unposted_final_score)" } else { '' }), $(if ($grades.current_grade) { $grades.current_grade } else { '' }), $(if ($grades.unposted_current_grade) { $grades.unposted_current_grade } else { '' }), $(if ($grades.final_grade) { $grades.final_grade } else { '' }), $(if ($grades.unposted_final_grade) { $grades.unposted_final_grade } else { '' }) )
                                foreach ($asgn in $assignments) { $subKey = $sid + "_" + $asgn.id; $row += if ($subLookup.ContainsKey($subKey)) { $subLookup[$subKey] } else { '' } }
                                $lines += (($row | ForEach-Object { Format-CsvField $_ }) -join ',')
                            }
                            $tmpCsv = Join-Path $TempDir "gradebook-$courseId-$([Guid]::NewGuid()).csv"
                            [System.IO.File]::WriteAllBytes($tmpCsv, ([System.Text.Encoding]::UTF8.GetPreamble() + [System.Text.Encoding]::UTF8.GetBytes($lines -join "`r`n")))
                            Publish-FileToS3 -LocalPath $tmpCsv -Key $c.DailyGrades
                            Invoke-TieredPromotion -BaseKey "$($c.BaseKey)-gradebook" -Ext '.csv'
                            Remove-TempPath -Path $tmpCsv
                            $output += [pscustomobject]@{ CourseId = $courseId; Action = 'written'; Items = $enrollments.Count }
                        } else { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'skipped' } }
                    } catch { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = $_.Exception.Message } }
                }

                elseif ($TaskType -eq 'Pages') {
                    if ($WhatIfMode) { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun' }; continue }
                    try {
                        $pages = @(Invoke-CanvasGet "/courses/$courseId/pages" @{ 'include[]' = 'body' })
                        if ($pages.Count -gt 0) {
                            $jsonText = $pages | ConvertTo-Json -Depth 10 -Compress
                            Publish-TextToS3 -Text $jsonText -Key $c.DailyPages
                            Invoke-TieredPromotion -BaseKey "$($c.BaseKey)-pages" -Ext '.json'
                            $output += [pscustomobject]@{ CourseId = $courseId; Action = 'written'; Items = $pages.Count }
                        } else { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'skipped' } }
                    } catch { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = $_.Exception.Message } }
                }
                
                elseif ($TaskType -eq 'Files') {
                    if ($WhatIfMode) { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun' }; continue }
                    try {
                        $folders = @(Invoke-CanvasGet "/courses/$courseId/folders")
                        $files = @(Invoke-CanvasGet "/courses/$courseId/files" @{ 'per_page' = '500' })
                        
                        if ($files.Count -gt 0) {
                            $totalCourseSize = 0
                            foreach ($f in $files) { $totalCourseSize += [long]$f.size }
                            
                            # Bin Packing limits adjusted to 10GB/20GB chunks to safely scale with high concurrency.
                            $SplitThreshold = 20GB
                            $MaxChunkSize = 10GB
                            $maxBlockSize = if ($totalCourseSize -gt $SplitThreshold) { $MaxChunkSize } else { $totalCourseSize + 1GB }

                            $folderMap = @{}; foreach ($f in $folders) { $folderMap["$($f.id)"] = $f.full_name }

                            $currentBlock = @()
                            $currentBlockSize = 0
                            $blockNumber = 1
                            $totalItemsProcessed = 0

                            function Process-FileBlock {
                                param($BlockFiles, $BNum)
                                if ($BlockFiles.Count -eq 0) { return }

                                $localDir = Join-Path $TempDir "files-$courseId-$([Guid]::NewGuid())"
                                $zipPath = Join-Path $TempDir "files-$courseId-part$BNum-$([Guid]::NewGuid()).zip"

                                try {
                                    # Safeguard: Block jobs if global cache exceeds 100GB
                                    $maxCacheBytes = 100GB
                                    while ($true) {
                                        $currentCacheBytes = 0
                                        if (Test-Path $TempDir) {
                                            $currentCacheBytes = (Get-ChildItem $TempDir -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                                        }
                                        if ($null -eq $currentCacheBytes) { $currentCacheBytes = 0 }
                                        if ($currentCacheBytes -lt $maxCacheBytes) { break }
                                        Start-Sleep -Seconds 10
                                    }

                                    New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                                    
                                    foreach ($bf in $BlockFiles) {
                                        if (-not $bf.url) { continue }
                                        $fId = "$($bf.folder_id)"
                                        $path = if ($folderMap.ContainsKey($fId)) { $folderMap[$fId] } else { "" }
                                        $path = $path -replace '(?i)^course files[/\\]?', ''

                                        $targetDir = Join-Path $localDir $path
                                        if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                                        
                                        try {
                                            Invoke-WebRequest -Uri $bf.url -OutFile (Join-Path $targetDir $bf.display_name) -UseBasicParsing -TimeoutSec 60
                                        } catch {
                                            # Gracefully skip dead links/ghost files so the zip doesn't abort
                                        }
                                    }

                                    # Using Native .NET ZipFile to bypass 2GB Compress-Archive bug
                                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                                    [System.IO.Compression.ZipFile]::CreateFromDirectory($localDir, $zipPath, [System.IO.Compression.CompressionLevel]::Fastest, $false)

                                    $blockKey = if ($totalCourseSize -gt $SplitThreshold) { $c.DailyFiles -replace '\.zip$', "-part$BNum.zip" } else { $c.DailyFiles }
                                    $blockExt = if ($totalCourseSize -gt $SplitThreshold) { "-part$BNum.zip" } else { ".zip" }

                                    Publish-FileToS3 -LocalPath $zipPath -Key $blockKey
                                    Invoke-TieredPromotion -BaseKey "$($c.BaseKey)-files" -Ext $blockExt

                                } finally {
                                    Remove-TempPath -Path $localDir -Recurse
                                    Remove-TempPath -Path $zipPath
                                }
                            }

                            foreach ($file in $files) {
                                $fSize = [long]$file.size
                                if (($currentBlockSize + $fSize) -gt $maxBlockSize -and $currentBlock.Count -gt 0) {
                                    Process-FileBlock -BlockFiles $currentBlock -BNum $blockNumber
                                    $totalItemsProcessed += $currentBlock.Count
                                    $currentBlock = @()
                                    $currentBlockSize = 0
                                    $blockNumber++
                                }
                                $currentBlock += $file
                                $currentBlockSize += $fSize
                            }

                            if ($currentBlock.Count -gt 0) {
                                Process-FileBlock -BlockFiles $currentBlock -BNum $blockNumber
                                $totalItemsProcessed += $currentBlock.Count
                            }

                            $output += [pscustomobject]@{ CourseId = $courseId; Action = 'written'; Items = $totalItemsProcessed }
                        } else { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'skipped' } }
                    } catch { $output += [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = $_.Exception.Message } }
                }
            }
            return $output
        }

        # ------------------------------------------------------------
        # Gradebook export pass (Parallelized)
        # ------------------------------------------------------------
        $gradebookResults = @()
        if ($SkipGradebooks) { Write-Warn "Skipping Gradebook export pass" }
        else {
            Write-Info "=== Backing up Gradebooks ($ConcurrentJobs Parallel Jobs) ==="
            $gradesToExport = @()
            foreach ($c in $exportList) {
                if (-not $Force -and $s3Cache.ContainsKey($c.DailyGrades)) { $gradebookResults += [pscustomobject]@{ CourseId = $c.CourseId; Action = 'skipped' } }
                else { $gradesToExport += $c }
            }
            
            $sliceSize = if ($gradesToExport.Count -gt 0) { [math]::Ceiling($gradesToExport.Count / $ConcurrentJobs) } else { 1 }
            $jobChunks = @()
            for ($i = 0; $i -lt $gradesToExport.Count; $i += $sliceSize) { $jobChunks += ,@($gradesToExport[$i..[math]::Min($i+$sliceSize-1, $gradesToExport.Count - 1)]) }

            foreach ($chunk in $jobChunks) {
                $argsHash = @{ TaskType = 'Gradebook'; Chunk = $chunk; CanvasApiToken = $CanvasApiToken; CanvasBaseUrl = $CanvasBaseUrl; S3Bucket = $S3Bucket; AwsRegion = $AwsRegion; TempDir = $TempDir; Force = $Force; WhatIfMode = $WhatIfMode; TodayString = $todayString }
                Start-Job -Name "GradebookJob" -ScriptBlock $DataWorkerJobBlock -ArgumentList $argsHash | Out-Null
            }
            Get-Job -Name "GradebookJob" | Wait-Job | ForEach-Object { $gradebookResults += Receive-Job -Job $_; Remove-Job -Job $_ }
        }
        $gW = @($gradebookResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count

        # ------------------------------------------------------------
        # Pages export pass (Parallelized)
        # ------------------------------------------------------------
        $pagesResults = @()
        if ($SkipPages) { Write-Warn "Skipping Pages export pass" }
        else {
            Write-Info "=== Backing up Pages ($ConcurrentJobs Parallel Jobs) ==="
            $pagesToExport = @()
            foreach ($c in $exportList) {
                if (-not $Force -and $s3Cache.ContainsKey($c.DailyPages)) { $pagesResults += [pscustomobject]@{ CourseId = $c.CourseId; Action = 'skipped' } }
                else { $pagesToExport += $c }
            }
            
            $sliceSize = if ($pagesToExport.Count -gt 0) { [math]::Ceiling($pagesToExport.Count / $ConcurrentJobs) } else { 1 }
            $jobChunks = @()
            for ($i = 0; $i -lt $pagesToExport.Count; $i += $sliceSize) { $jobChunks += ,@($pagesToExport[$i..[math]::Min($i+$sliceSize-1, $pagesToExport.Count - 1)]) }

            foreach ($chunk in $jobChunks) {
                $argsHash = @{ TaskType = 'Pages'; Chunk = $chunk; CanvasApiToken = $CanvasApiToken; CanvasBaseUrl = $CanvasBaseUrl; S3Bucket = $S3Bucket; AwsRegion = $AwsRegion; TempDir = $TempDir; Force = $Force; WhatIfMode = $WhatIfMode; TodayString = $todayString }
                Start-Job -Name "PagesJob" -ScriptBlock $DataWorkerJobBlock -ArgumentList $argsHash | Out-Null
            }
            Get-Job -Name "PagesJob" | Wait-Job | ForEach-Object { $pagesResults += Receive-Job -Job $_; Remove-Job -Job $_ }
        }
        $pW = @($pagesResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count

        # ------------------------------------------------------------
        # Files export pass (Parallelized)
        # ------------------------------------------------------------
        $filesResults = @()
        if ($SkipFiles) { Write-Warn "Skipping Files export pass" }
        else {
            Write-Info "=== Backing up Files ($ConcurrentJobs Parallel Jobs) ==="
            $filesToExport = @()
            foreach ($c in $exportList) {
                if (-not $Force -and $s3Cache.ContainsKey($c.DailyFiles)) { $filesResults += [pscustomobject]@{ CourseId = $c.CourseId; Action = 'skipped' } }
                else { $filesToExport += $c }
            }
            
            $sliceSize = if ($filesToExport.Count -gt 0) { [math]::Ceiling($filesToExport.Count / $ConcurrentJobs) } else { 1 }
            $jobChunks = @()
            for ($i = 0; $i -lt $filesToExport.Count; $i += $sliceSize) { $jobChunks += ,@($filesToExport[$i..[math]::Min($i+$sliceSize-1, $filesToExport.Count - 1)]) }

            foreach ($chunk in $jobChunks) {
                $argsHash = @{ TaskType = 'Files'; Chunk = $chunk; CanvasApiToken = $CanvasApiToken; CanvasBaseUrl = $CanvasBaseUrl; S3Bucket = $S3Bucket; AwsRegion = $AwsRegion; TempDir = $TempDir; Force = $Force; WhatIfMode = $WhatIfMode; TodayString = $todayString }
                Start-Job -Name "FilesJob" -ScriptBlock $DataWorkerJobBlock -ArgumentList $argsHash | Out-Null
            }
            Get-Job -Name "FilesJob" | Wait-Job | ForEach-Object { $filesResults += Receive-Job -Job $_; Remove-Job -Job $_ }
        }
        $fW = @($filesResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count

        # ------------------------------------------------------------
        # Sub-Account Pass
        # ------------------------------------------------------------
        if (-not $SkipSubAccounts) {
            Write-Info "=== Backing up all sub-accounts ==="
            try {
                $subAccounts = @(Get-AllSubAccountsRecursive -RootId $RootAccountId)
                foreach ($sub in $subAccounts) { try { Backup-SubAccount -AccountId $sub.id } catch { Write-Fail "Sub-account fail: $_" } }
            } catch { Write-Fail "Enumerate sub-accounts failed: $_" }
        }

        # ------------------------------------------------------------
        # Coverage Verification
        # ------------------------------------------------------------
        if ($CourseId) {
            $coverage = [pscustomobject]@{ CoveragePercent = 100; CanvasCourses = 1; S3Courses = 1 }
        } else {
            $semester = switch ($today.Month) { { $_ -in 1..5 } { 'spring' }; { $_ -in 6..8 } { 'summer' }; default { 'fall' } }
            $coverage = Test-BackupCoverage -TermIds $termIds -Year $thisYear.ToString() -Semester $semester
        }

        $duration = (Get-Date) - $startTime
        Write-Info ''
        Write-Info '=== SUMMARY ==='
        Write-Info "Duration : $([int]$duration.TotalMinutes) min $($duration.Seconds) sec"
        Write-Info "Content (.imscc)  : $cW written"
        Write-Info "Gradebook (.csv)  : $gW written"
        Write-Info "Pages (.json)     : $pW written"
        Write-Info "Files (.zip)      : $fW written"
        Write-Info "coverage_percent=$($coverage.CoveragePercent)  canvas_courses=$($coverage.CanvasCourses)  s3_courses=$($coverage.S3Courses)"

        if ($WhatIfMode) { $runStatus = 'DRYRUN'; Write-Warn 'BACKUP DRY RUN COMPLETE' }
        elseif ($coverage.CoveragePercent -lt 90) { $runStatus = 'FAILED'; Write-Fail 'BACKUP FAILURE' }
        else { $runStatus = 'COMPLETE'; Write-Ok 'BACKUP COMPLETE' }

        $logs = @(Get-ChildItem $LogDir -Filter '*.log' | Sort-Object LastWriteTime -Descending)
        if ($logs.Count -gt 30) { $logs | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue }
    }
    finally {
        if ($transcriptStarted) { Stop-Transcript | Out-Null }

        $endTime = Get-Date
        $duration = $endTime - $startTime
        $summaryText = @(
            "run_status=$runStatus"
            "start_time=$($startTime.ToString('s'))"
            "end_time=$($endTime.ToString('s'))"
            "duration_minutes=$([int]$duration.TotalMinutes)"
            "content_written=$cW"
            "gradebook_written=$gW"
            "pages_written=$pW"
            "files_written=$fW"
            "coverage_percent=$(if ($coverage) { $coverage.CoveragePercent } else { '' })"
            "canvas_courses=$(if ($coverage) { $coverage.CanvasCourses } else { '' })"
            "s3_courses=$(if ($coverage) { $coverage.S3Courses } else { '' })"
            "s3_log_key=$s3LogKey"
        ) -join "`r`n"

        try { if (Test-Path $transcriptPath) { Publish-FileToS3 -LocalPath $transcriptPath -Key $s3LogKey -StorageClass 'STANDARD' } } catch {}
        try { Publish-TextToS3 -Text $summaryText -Key $s3SummaryKey } catch {}
        
        # Final cleanup sweep to ensure empty Temp directory between runs
        Remove-Item -Path "$TempDir\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
# Script body - resolve configuration, then run
# ============================================================
$config = Import-CanvasBackupConfig

if (-not $CanvasBaseUrl -and $config)   { $CanvasBaseUrl   = $config.CanvasBaseUrl }
if (-not $CanvasApiToken -and $config)  { $CanvasApiToken  = ConvertFrom-SecureStringToPlainText $config.CanvasApiToken }
if (-not $S3Bucket -and $config)        { $S3Bucket        = $config.S3Bucket }
if (-not $AwsRegion -and $config)       { $AwsRegion       = $config.AwsRegion }
if (-not $RootAccountId -and $config)   { $RootAccountId   = $config.RootAccountId }

if (-not $CanvasBaseUrl)  { $CanvasBaseUrl  = $env:CANVAS_BASE_URL }
if (-not $CanvasApiToken) { $CanvasApiToken = $env:CANVAS_API_TOKEN }
if (-not $S3Bucket)       { $S3Bucket       = $env:S3_BUCKET }

if (-not $AwsRegion) {
    if ($env:S3_REGION) { $AwsRegion = $env:S3_REGION }
    elseif ($env:AWS_REGION) { $AwsRegion = $env:AWS_REGION }
    else { $AwsRegion = 'us-east-1' }
}
if (-not $RootAccountId) { if ($env:CANVAS_ACCOUNT_ID) { $RootAccountId = $env:CANVAS_ACCOUNT_ID } else { $RootAccountId = '1' } }

Main