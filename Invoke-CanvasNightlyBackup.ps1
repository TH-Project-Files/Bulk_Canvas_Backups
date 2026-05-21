#Requires -Version 5.1
<#
.SYNOPSIS
    Nightly Canvas LMS backup - self-contained PowerShell, no NPM/Node/LLM required.

.DESCRIPTION
    This script performs a nightly backup workflow for Canvas content and gradebooks.

    High-level flow:
      1. Load encrypted configuration from C:\Scripts\canvas-backup-config.clixml
      2. In normal mode:
           - resolve active enrollment terms from current + prior calendar year
           - enumerate active-term courses from the Canvas root account
      3. In single-course test mode (-CourseId):
           - fetch only the specified course directly from Canvas
           - skip root-account term enumeration and global course listing
      4. Back up course content packages and upload them to S3
      5. Back up gradebook CSVs and upload them to S3
      6. Optionally recurse through all Canvas sub-accounts for persistent backups
      7. Verify S3 coverage against Canvas course counts
           - skipped in single-course mode
      8. Write a local transcript log and upload both the log and a summary file to S3

    Hard dependencies:
      - PowerShell 5.1+ (ships with Windows; no install needed)
      - AWS CLI v2 (aws.exe)
      - Network access to the Canvas REST API and AWS S3

    FERPA note:
      Gradebook CSVs contain student grade data. This script writes only to the configured
      S3 bucket. Ensure bucket encryption and IAM restrictions are already in place.

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
    Defaults to C:\Scripts\Temp.

.PARAMETER LogDir
    Directory for transcript log files.
    Defaults to C:\Scripts\Logs.

.PARAMETER PollTimeoutMins
    Maximum minutes to wait for a single Canvas export to complete. Default: 60.

.PARAMETER PollIntervalSecs
    Seconds between Canvas export status polls. Default: 3.

.PARAMETER CourseId
    Optional single-course mode. If provided, only this Canvas course ID is processed
    in the main backup pass, and root-account term/course discovery is skipped.

.PARAMETER MaxCourses
    Optional test-mode limiter. If greater than 0, only the first N active-term
    courses are processed in the main active-term backup pass.

.PARAMETER SkipSubAccounts
    Skip the recursive sub-account backup pass.

.PARAMETER SkipContent
    Skip course content export backups.

.PARAMETER SkipGradebooks
    Skip gradebook CSV backups.

.PARAMETER WhatIfMode
    Dry-run switch. No Canvas exports, S3 uploads, S3 copies, or S3 deletes occur.
    The script still resolves terms/courses and logs what it would do.

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
    [string]$CourseId,
    [int]$MaxCourses = 0,
    [switch]$SkipSubAccounts,
    [switch]$SkipContent,
    [switch]$SkipGradebooks,
    [switch]$WhatIfMode,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
# Canvas API helpers
# ============================================================
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
            if ($response -and $response.StatusCode) {
                $code = [int]$response.StatusCode
            }

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
            else {
                throw
            }
        }
    }
}

function Invoke-CanvasGet {
    param(
        [string]$Endpoint,
        [hashtable]$Query = @{}
    )

    $base = $CanvasBaseUrl.TrimEnd('/')
    $qs = @('per_page=100')
    foreach ($kv in $Query.GetEnumerator()) {
        $qs += "$([Uri]::EscapeDataString($kv.Key))=$([Uri]::EscapeDataString([string]$kv.Value))"
    }

    $url = "$base/api/v1${Endpoint}?" + ($qs -join '&')
    $results = @()

    do {
        $resp = Invoke-CanvasApiPage -Url $url
        $page = $resp.Content | ConvertFrom-Json
        $results += @($page)

        $url = $null
        $linkHeader = $resp.Headers['Link']
        if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) {
            $url = $Matches[1]
        }
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
        if ($page.enrollment_terms) { $results += @($page.enrollment_terms) }

        $url = $null
        $linkHeader = $resp.Headers['Link']
        if ($linkHeader -and ($linkHeader -match '<([^>]+)>;\s*rel="next"')) {
            $url = $Matches[1]
        }
    } while ($url)

    return $results
}

function Invoke-CanvasPost {
    param(
        [string]$Endpoint,
        [hashtable]$Body = @{}
    )

    $uri = "$($CanvasBaseUrl.TrimEnd('/'))/api/v1${Endpoint}"
    $json = $Body | ConvertTo-Json -Compress

    $resp = Invoke-WebRequest -Uri $uri -Method Post `
        -Headers @{ Authorization = "Bearer $CanvasApiToken" } `
        -ContentType 'application/json' `
        -Body $json -UseBasicParsing -TimeoutSec 60

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
        $code = $null
        if ($response -and $response.StatusCode) {
            $code = [int]$response.StatusCode
        }

        if ($code -eq 404) {
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

    $json = & aws s3api list-objects-v2 `
        --bucket $S3Bucket `
        --prefix $Key `
        --max-items 1 `
        --output json `
        --region $AwsRegion 2>$null

    if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') {
        return $false
    }

    $page = $json | ConvertFrom-Json
    if ($null -eq $page) {
        return $false
    }

    $hasContents = $page.PSObject.Properties.Name -contains 'Contents'
    if (-not $hasContents -or $null -eq $page.Contents) {
        return $false
    }

    return @($page.Contents | Where-Object { $_.Key -eq $Key }).Count -gt 0
}

function Publish-FileToS3 {
    param(
        [string]$LocalPath,
        [string]$Key,
        [string]$StorageClass = 'GLACIER_IR'
    )

    if ($WhatIfMode) {
        Write-DryRun "Would upload '$LocalPath' to s3://$S3Bucket/$Key (storage-class=$StorageClass)"
        return
    }

    & aws s3 cp $LocalPath "s3://$S3Bucket/$Key" `
        --storage-class $StorageClass --region $AwsRegion --no-progress

    if ($LASTEXITCODE -ne 0) {
        throw "aws s3 cp failed for key: $Key"
    }
}

function Publish-TextToS3 {
    param(
        [string]$Text,
        [string]$Key
    )

    $tmpFile = Join-Path $TempDir "s3-text-$([Guid]::NewGuid()).txt"
    try {
        [System.IO.File]::WriteAllText($tmpFile, $Text, [System.Text.Encoding]::UTF8)
        Publish-FileToS3 -LocalPath $tmpFile -Key $Key -StorageClass 'STANDARD'
    }
    finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-S3CopyObject {
    param([string]$SourceKey, [string]$DestKey)

    if ($WhatIfMode) {
        Write-DryRun "Would copy s3://$S3Bucket/$SourceKey -> s3://$S3Bucket/$DestKey"
        return
    }

    & aws s3 cp "s3://$S3Bucket/$SourceKey" "s3://$S3Bucket/$DestKey" `
        --metadata-directive COPY --storage-class GLACIER_IR --region $AwsRegion --no-progress

    if ($LASTEXITCODE -ne 0) {
        throw "aws s3 cp (copy) failed: $SourceKey -> $DestKey"
    }
}

function Get-S3ObjectsUnderPrefix {
    param([string]$Prefix)

    $results = @()
    $token = $null

    do {
        $args = @(
            's3api', 'list-objects-v2',
            '--bucket', $S3Bucket,
            '--prefix', $Prefix,
            '--output', 'json',
            '--region', $AwsRegion
        )
        if ($token) {
            $args += @('--continuation-token', $token)
        }

        $json = & aws @args 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }

        $page = $json | ConvertFrom-Json
        if ($null -eq $page) { break }

        $hasContents = $page.PSObject.Properties.Name -contains 'Contents'
        if ($hasContents -and $null -ne $page.Contents) {
            $results += @($page.Contents | Select-Object Key, LastModified)
        }

        $hasIsTruncated = $page.PSObject.Properties.Name -contains 'IsTruncated'
        $hasNextToken   = $page.PSObject.Properties.Name -contains 'NextContinuationToken'

        if ($hasIsTruncated -and $page.IsTruncated -and $hasNextToken -and $page.NextContinuationToken) {
            $token = $page.NextContinuationToken
        }
        else {
            $token = $null
        }
    } while ($token)

    if (-not $results) { return @() }
    return @($results | Sort-Object { [datetime]$_.LastModified } -Descending)
}

function Get-S3KeysUnderPrefix {
    param([string]$Prefix)

    $results = @()
    $token = $null

    do {
        $args = @(
            's3api', 'list-objects-v2',
            '--bucket', $S3Bucket,
            '--prefix', $Prefix,
            '--output', 'json',
            '--region', $AwsRegion
        )
        if ($token) {
            $args += @('--continuation-token', $token)
        }

        $json = & aws @args 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $json -or $json -eq 'null') { break }

        $page = $json | ConvertFrom-Json
        if ($null -eq $page) { break }

        $hasContents = $page.PSObject.Properties.Name -contains 'Contents'
        if ($hasContents -and $null -ne $page.Contents) {
            $results += @($page.Contents | ForEach-Object { $_.Key })
        }

        $hasIsTruncated = $page.PSObject.Properties.Name -contains 'IsTruncated'
        $hasNextToken   = $page.PSObject.Properties.Name -contains 'NextContinuationToken'

        if ($hasIsTruncated -and $page.IsTruncated -and $hasNextToken -and $page.NextContinuationToken) {
            $token = $page.NextContinuationToken
        }
        else {
            $token = $null
        }
    } while ($token)

    return @($results)
}

function Invoke-S3Prune {
    param([string]$Prefix, [int]$KeepCount, [string]$Ext)

    $objects = @(Get-S3ObjectsUnderPrefix -Prefix $Prefix | Where-Object { $_.Key.EndsWith($Ext) })
    if ($objects.Count -le $KeepCount) { return }

    $toDelete = @($objects | Select-Object -Skip $KeepCount)
    if ($toDelete.Count -ge $objects.Count) {
        Write-Warn "Skipping prune under '$Prefix' - would delete all copies"
        return
    }

    foreach ($obj in $toDelete) {
        if ($WhatIfMode) {
            Write-DryRun "Would delete s3://$S3Bucket/$($obj.Key)"
        }
        else {
            & aws s3api delete-object --bucket $S3Bucket --key $obj.Key --region $AwsRegion | Out-Null
            Write-Info "Pruned: $($obj.Key)"
        }
    }
}

function Invoke-TieredPromotion {
    param([string]$BaseKey, [string]$Ext, [string]$Today)

    $dailyKey = "$BaseKey/daily/$Today$Ext"
    $weeklyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/weekly/")
    $monthlyObjs = @(Get-S3ObjectsUnderPrefix "$BaseKey/monthly/")

    $daysSinceWeekly = 999
    $daysSinceMonthly = 999

    if ($weeklyObjs.Count -gt 0) {
        $daysSinceWeekly = ((Get-Date) - [datetime]$weeklyObjs[0].LastModified).TotalDays
    }
    if ($monthlyObjs.Count -gt 0) {
        $daysSinceMonthly = ((Get-Date) - [datetime]$monthlyObjs[0].LastModified).TotalDays
    }

    $needsWeekly = $daysSinceWeekly -ge 7
    $needsMonthly = $daysSinceMonthly -ge 30

    if ($needsWeekly) {
        Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/weekly/$Today$Ext"
        Write-Info "Promoted to weekly: $BaseKey/weekly/$Today$Ext"
    }
    if ($needsWeekly -and $needsMonthly) {
        Invoke-S3CopyObject -SourceKey $dailyKey -DestKey "$BaseKey/monthly/$Today$Ext"
        Write-Info "Promoted to monthly: $BaseKey/monthly/$Today$Ext"
    }

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
            $season = $Matches[1].ToLower()
            $year = $Matches[2]
        }
        elseif ($TermName -match '(\d{4}).*?\b(fall|spring|summer|winter)\b') {
            $year = $Matches[1]
            $season = $Matches[2].ToLower()
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
        }
        catch {}
    }

    if (-not $year) {
        if ($TermName -and $TermName -match '(\d{4})') {
            $year = $Matches[1]
        }
        else {
            $year = (Get-Date).Year.ToString()
        }
    }

    if (-not $season) { $season = 'unknown' }

    [pscustomobject]@{
        Year = $year
        Semester = $season
    }
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

    $identifier = if ($Course.sis_course_id) {
        $Course.sis_course_id
    }
    else {
        $nameSlug = ConvertTo-CanvasSlug $Course.name
        "$nameSlug-$($TermParts.Semester)-$($TermParts.Year)-$($Course.id)"
    }

    "$BasePrefix/$($TermParts.Year)/$($TermParts.Semester)/$DeptSlug/$identifier"
}

# ============================================================
# Course content backup
# ============================================================
function Backup-CourseContent {
    param($Course, $TermParts, [string]$DeptSlug, [string]$BasePrefix = 'canvas-backups')

    $courseId = $Course.id
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $ext = '.imscc'
    $baseKey = Get-CourseS3BaseKey -Course $Course -TermParts $TermParts -DeptSlug $DeptSlug -BasePrefix $BasePrefix
    $dailyKey = "$baseKey/daily/$today$ext"

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Would submit Canvas content export and upload to s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }

    if (-not $Force -and (Test-S3KeyExists -Key $dailyKey)) {
        Write-Info "[$courseId] Skip - already backed up today"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Submitting export..."
    try {
        $job = Invoke-CanvasPost "/courses/$courseId/content_exports" @{
            export_type = 'common_cartridge'
            skip_notifications = $true
        }
    }
    catch {
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

        try {
            $exp = (Invoke-CanvasGet "/courses/$courseId/content_exports/$($job.id)" | Select-Object -First 1)
        }
        catch {
            Write-Warn "[$courseId] Poll error (will retry): $_"
            $exp = $null
            continue
        }

        if ($exp -and $exp.workflow_state -eq 'failed') {
            Write-Fail "[$courseId] Canvas export failed"
            return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Canvas export failed' }
        }

        if ($exp) {
            Write-Info "[$courseId] State: $($exp.workflow_state)"
        }
    } until ($exp -and $exp.workflow_state -eq 'exported')

    if (-not $exp.attachment -or -not $exp.attachment.url) {
        Write-Fail "[$courseId] Export completed but no attachment URL was returned"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = 'Missing attachment URL' }
    }

    $downloadUrl = $exp.attachment.url
    $tempFile = Join-Path $TempDir "canvas-export-$courseId-$([Guid]::NewGuid()).imscc"

    try {
        Write-Info "[$courseId] Downloading..."
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 600
        $sizeBytes = (Get-Item $tempFile).Length

        Write-Info "[$courseId] Uploading -> $dailyKey"
        Publish-FileToS3 -LocalPath $tempFile -Key $dailyKey
        Invoke-TieredPromotion -BaseKey $baseKey -Ext $ext -Today $today

        Write-Ok "[$courseId] $([Math]::Round($sizeBytes / 1MB, 1)) MB -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{
            CourseId = $courseId
            Action = 'written'
            S3Key = $dailyKey
            SizeBytes = $sizeBytes
        }
    }
    catch {
        Write-Fail "[$courseId] Upload/download error: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }
    finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
    }
}

# ============================================================
# Gradebook backup
# ============================================================
function Format-CsvField {
    param([string]$Value)

    if ($null -eq $Value) { $Value = '' }
    '"' + ($Value -replace '"', '""') + '"'
}

function Backup-CourseGradebook {
    param($Course, $TermParts, [string]$DeptSlug, [string]$BasePrefix = 'canvas-backups')

    $courseId = $Course.id
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $ext = '.csv'
    $baseKey = Get-CourseS3BaseKey -Course $Course -TermParts $TermParts -DeptSlug $DeptSlug -BasePrefix $BasePrefix
    $gbBaseKey = "$baseKey-gradebook"
    $dailyKey = "$gbBaseKey/daily/$today$ext"

    if ($WhatIfMode) {
        Write-DryRun "[$courseId] Would fetch gradebook data and upload CSV to s3://$S3Bucket/$dailyKey"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'dryrun'; S3Key = $dailyKey }
    }

    if (-not $Force -and (Test-S3KeyExists -Key $dailyKey)) {
        Write-Info "[$courseId] Gradebook skip - already written today"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; S3Key = $dailyKey }
    }

    Write-Info "[$courseId] Fetching gradebook data..."
    try {
        $assignments = @(Invoke-CanvasGet "/courses/$courseId/assignments")
        $enrollments = @(Invoke-CanvasGet "/courses/$courseId/enrollments" @{
            'type[]' = 'StudentEnrollment'
            'include[]' = 'user,grades'
        })
        $submissions = @(Invoke-CanvasGet "/courses/$courseId/students/submissions" @{
            'student_ids[]' = 'all'
            'per_page' = '100'
        })
        $asgnGroups = @(Invoke-CanvasGet "/courses/$courseId/assignment_groups" @{
            'include[]' = 'assignments'
        })
    }
    catch {
        Write-Fail "[$courseId] Gradebook fetch failed: $_"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }

    if ($enrollments.Count -eq 0) {
        Write-Warn "[$courseId] No student enrollments - skipping gradebook"
        return [pscustomobject]@{ CourseId = $courseId; Action = 'skipped'; SkipReason = 'no students' }
    }

    $subLookup = @{}
    foreach ($sub in $submissions) {
        $key = ($sub.user_id.ToString()) + "_" + ($sub.assignment_id.ToString())
        $subLookup[$key] = if ($null -ne $sub.score) { "$($sub.score)" } else { '' }
    }

    $groupNames = @{}
    foreach ($grp in $asgnGroups) {
        $groupNames["$($grp.id)"] = $grp.name
    }

    $header = @(
        'Student', 'ID', 'SIS User ID', 'SIS Login ID', 'Section',
        'Current Score', 'Unposted Current Score',
        'Final Score', 'Unposted Final Score',
        'Current Grade', 'Unposted Current Grade',
        'Final Grade', 'Unposted Final Grade'
    )

    foreach ($asgn in $assignments) {
        $gName = if ($groupNames.ContainsKey("$($asgn.assignment_group_id)")) {
            $groupNames["$($asgn.assignment_group_id)"]
        }
        else { '' }

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
        Invoke-TieredPromotion -BaseKey $gbBaseKey -Ext $ext -Today $today

        Write-Ok "[$courseId] Gradebook ($($enrollments.Count) students) -> s3://$S3Bucket/$dailyKey"
        [pscustomobject]@{
            CourseId = $courseId
            Action = 'written'
            S3Key = $dailyKey
            Students = $enrollments.Count
            Assignments = $assignments.Count
        }
    }
    catch {
        Write-Fail "[$courseId] Gradebook upload failed: $_"
        [pscustomobject]@{ CourseId = $courseId; Action = 'failed'; Error = "$_" }
    }
    finally {
        if (Test-Path $tmpCsv) { Remove-Item $tmpCsv -Force -ErrorAction SilentlyContinue }
    }
}

# ============================================================
# Persistent sub-account backup
# ============================================================
function Backup-SubAccount {
    param([string]$AccountId)

    Write-Info "=== Sub-account backup: account $AccountId ==="
    $accountInfo = (Invoke-CanvasGet "/accounts/${AccountId}" | Select-Object -First 1)
    $accountSlug = ConvertTo-CanvasSlug $accountInfo.name
    if (-not $accountSlug) { $accountSlug = "account-$AccountId" }
    $basePrefix = "canvas-backups/persistent/$accountSlug"

    $courses = @(Invoke-CanvasGet "/accounts/${AccountId}/courses" @{
        'include[]' = 'term,account_name'
        'per_page' = '100'
    })
    $courses = @($courses | Group-Object id | ForEach-Object { $_.Group[0] })

    Write-Info "Sub-account '$($accountInfo.name)': $($courses.Count) courses"

    $contentResults = @()
    $gradebookResults = @()

    foreach ($course in $courses) {
        $termName = if ($course.term) { $course.term.name } else { '' }
        $startAt = if ($course.term) { $course.term.start_at } else { '' }
        $termParts = Get-TermParts -TermName $termName -StartAt $startAt
        $acctSlug = if ($course.account_name) { ConvertTo-CanvasSlug $course.account_name } else { '' }
        $deptSlug = Get-DeptSlug -AccountSlug $acctSlug -CourseCode $course.course_code

        $contentResults += Backup-CourseContent -Course $course -TermParts $termParts -DeptSlug $deptSlug -BasePrefix $basePrefix
        $gradebookResults += Backup-CourseGradebook -Course $course -TermParts $termParts -DeptSlug $deptSlug -BasePrefix $basePrefix
    }

    $written = @($contentResults | Where-Object { $_.Action -eq 'written' }).Count
    $skipped = @($contentResults | Where-Object { $_.Action -eq 'skipped' }).Count
    $failed  = @($contentResults | Where-Object { $_.Action -eq 'failed' }).Count
    $dryrun  = @($contentResults | Where-Object { $_.Action -eq 'dryrun' }).Count

    Write-Ok "Sub-account '$($accountInfo.name)' done: $written written, $skipped skipped, $failed failed, $dryrun dryrun"
}

# ============================================================
# Coverage verification
# ============================================================
function Test-BackupCoverage {
    param([string[]]$TermIds, [string]$Year, [string]$Semester)

    Write-Info "=== Verifying coverage for $Year/$Semester ==="

    $canvasCourses = @()
    foreach ($tid in $TermIds) {
        $canvasCourses += @(Invoke-CanvasGet "/accounts/${RootAccountId}/courses" @{
            enrollment_term_id = $tid
        })
    }

    $canvasCourses = @($canvasCourses | Group-Object id | ForEach-Object { $_.Group[0] })
    $canvasCount = $canvasCourses.Count

    $prefix = "canvas-backups/$Year/$Semester/"
    $keys = @(Get-S3KeysUnderPrefix -Prefix $prefix)
    $s3Count = 0

    if ($keys.Count -gt 0) {
        $unique = $keys | ForEach-Object {
            $parts = $_ -split '/'
            if ($parts.Length -ge 5) {
                ($parts[0..4] -join '/')
            }
        } | Where-Object { $_ } | Sort-Object -Unique

        $s3Count = @($unique).Count
    }

    $coveragePct = if ($canvasCount -gt 0) {
        [int]([Math]::Round(($s3Count / $canvasCount) * 100))
    }
    else {
        0
    }

    if ($coveragePct -lt 90) {
        Write-Fail "BACKUP FAILURE: coverage_percent=$coveragePct  canvas_courses=$canvasCount  s3_courses=$s3Count"
    }
    else {
        Write-Ok "BACKUP OK:      coverage_percent=$coveragePct  canvas_courses=$canvasCount  s3_courses=$s3Count"
    }

    [pscustomobject]@{
        CoveragePercent = $coveragePct
        CanvasCourses = $canvasCount
        S3Courses = $s3Count
    }
}

# ============================================================
# Main orchestration
# ============================================================
function Main {
    $runStamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $yearMonth = Get-Date -Format 'yyyy/MM'
    $logFileName = "backup-$runStamp.log"
    $summaryFileName = "backup-$runStamp.summary.txt"
    $transcriptPath = Join-Path $LogDir $logFileName
    $s3LogKey = "canvas-backups/logs/$yearMonth/$logFileName"
    $s3SummaryKey = "canvas-backups/logs/$yearMonth/$summaryFileName"

    $transcriptStarted = $false
    $cW = 0; $cS = 0; $cF = 0; $cD = 0
    $gW = 0; $gS = 0; $gF = 0; $gD = 0
    $coverage = $null
    $runStatus = 'FAILED'
    $startTime = Get-Date

    $today = Get-Date
    $thisYear = $today.Year
    $prevYear = $today.Year - 1
    $accountCache = @{}
    $allCourses = @()
    $activeTerms = @()
    $termIds = @()

    if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
        throw "AWS CLI (aws.exe) not found in PATH."
    }
    if (-not $CanvasBaseUrl) { throw "CANVAS_BASE_URL is required" }
    if (-not $CanvasApiToken) { throw "CANVAS_API_TOKEN is required" }
    if (-not $S3Bucket) { throw "S3_BUCKET is required" }

    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
    if (-not (Test-Path $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    try {
        Start-Transcript -Path $transcriptPath | Out-Null
        $transcriptStarted = $true

        Write-Info "=== Canvas Nightly Backup starting at $startTime ==="
        Write-Info "Canvas: $CanvasBaseUrl"
        Write-Info "S3: s3://$S3Bucket (region: $AwsRegion)"
        Write-Info "Root Account ID: $RootAccountId"
        Write-Info "Local LogDir: $LogDir"
        Write-Info "Local TempDir: $TempDir"
        Write-Info "S3 Log Key: $s3LogKey"

        if ($Force) { Write-Warn "Force mode - same-day dedup bypassed" }
        if ($CourseId) { Write-Warn "Test mode: filtering to CourseId=$CourseId" }
        if ($MaxCourses -gt 0) { Write-Warn "Test mode: limiting to first $MaxCourses course(s)" }
        if ($SkipSubAccounts) { Write-Warn "Test mode: SkipSubAccounts enabled" }
        if ($SkipContent) { Write-Warn "Test mode: SkipContent enabled" }
        if ($SkipGradebooks) { Write-Warn "Test mode: SkipGradebooks enabled" }
        if ($WhatIfMode) { Write-Warn "DRY RUN mode enabled - no Canvas exports, S3 uploads, S3 copies, or S3 deletes will occur" }

        if ($CourseId) {
            Write-Info "=== Single-course mode ==="
            $singleCourse = Get-CanvasCourseById -CourseId $CourseId

            if (-not $singleCourse) {
                Write-Warn "CourseId $CourseId was not found"
                $runStatus = 'NO_MATCHING_COURSES'
                return
            }

            $allCourses = @($singleCourse)
            Write-Info "Single-course mode enabled: CourseId=$CourseId  matched=1"
        }
        else {
            Write-Info "=== Resolving enrollment terms ==="
            $allTerms = @(Get-EnrollmentTerms -AccountId $RootAccountId)
            $activeTerms = @(
                $allTerms | Where-Object {
                    $_.start_at -and ([datetime]$_.start_at).Year -in @($thisYear, $prevYear)
                }
            )
            $termIds = @($activeTerms | ForEach-Object { "$($_.id)" })
            Write-Info "Active terms ($($activeTerms.Count)): $($termIds -join ', ')"

            if ($activeTerms.Count -eq 0) {
                Write-Warn "No active terms found - nothing to back up"
                $runStatus = 'NO_ACTIVE_TERMS'
                return
            }

            Write-Info "=== Listing courses ==="
            foreach ($term in $activeTerms) {
                $tc = @(Invoke-CanvasGet "/accounts/${RootAccountId}/courses" @{
                    enrollment_term_id = $term.id
                    'include[]' = 'term,account_name'
                })
                Write-Info "  '$($term.name)': $($tc.Count) courses"
                $allCourses += $tc
            }

            $allCourses = @($allCourses | Group-Object id | ForEach-Object { $_.Group[0] })
            Write-Info "Total unique courses before test filters: $($allCourses.Count)"

            if ($MaxCourses -gt 0) {
                $allCourses = @($allCourses | Select-Object -First $MaxCourses)
                Write-Info "MaxCourses mode enabled: processing first $($allCourses.Count) courses"
            }

            if ($allCourses.Count -eq 0) {
                Write-Warn "No courses matched the requested filter"
                $runStatus = 'NO_MATCHING_COURSES'
                return
            }

            Write-Info "Total unique courses after test filters: $($allCourses.Count)"
        }

        $resolveTermAndDept = {
            param($c)

            $termObj = $null
            if ($activeTerms.Count -gt 0) {
                $termObj = $activeTerms | Where-Object { $_.id -eq $c.enrollment_term_id } | Select-Object -First 1
            }

            $termName = if ($termObj) { $termObj.name } elseif ($c.term) { $c.term.name } else { '' }
            $startAt  = if ($termObj) { $termObj.start_at } elseif ($c.term) { $c.term.start_at } else { '' }

            $tp = Get-TermParts -TermName $termName -StartAt $startAt

            $acctKey = "$($c.account_id)"
            if ($acctKey -and -not $accountCache.ContainsKey($acctKey)) {
                try {
                    $acct = (Invoke-CanvasGet "/accounts/${acctKey}" | Select-Object -First 1)
                    $accountCache[$acctKey] = ConvertTo-CanvasSlug $acct.name
                }
                catch {
                    $accountCache[$acctKey] = 'other'
                }
            }

            $acctSlug = $accountCache[$acctKey]
            $dept = Get-DeptSlug -AccountSlug $acctSlug -CourseCode $c.course_code

            @{
                TermParts = $tp
                DeptSlug = $dept
            }
        }

        $contentResults = @()
        if ($SkipContent) {
            Write-Warn "Skipping course content export pass"
        }
        else {
            Write-Info "=== Backing up course content exports ==="
            foreach ($c in $allCourses) {
                $info = & $resolveTermAndDept $c
                $contentResults += Backup-CourseContent -Course $c -TermParts $info.TermParts -DeptSlug $info.DeptSlug
            }
        }

        $cW = @($contentResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count
        $cS = @($contentResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'skipped' }).Count
        $cF = @($contentResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'failed' }).Count
        $cD = @($contentResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'dryrun' }).Count
        Write-Info "Content exports: written=$cW  skipped=$cS  failed=$cF  dryrun=$cD"

        $gradebookResults = @()
        if ($SkipGradebooks) {
            Write-Warn "Skipping gradebook export pass"
        }
        else {
            Write-Info "=== Backing up gradebooks ==="
            foreach ($c in $allCourses) {
                $info = & $resolveTermAndDept $c
                $gradebookResults += Backup-CourseGradebook -Course $c -TermParts $info.TermParts -DeptSlug $info.DeptSlug
            }
        }

        $gW = @($gradebookResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'written' }).Count
        $gS = @($gradebookResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'skipped' }).Count
        $gF = @($gradebookResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'failed' }).Count
        $gD = @($gradebookResults | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'dryrun' }).Count
        Write-Info "Gradebooks: written=$gW  skipped=$gS  failed=$gF  dryrun=$gD"

        if ($SkipSubAccounts) {
            Write-Warn "Skipping recursive sub-account backup pass"
        }
        else {
            Write-Info "=== Backing up all sub-accounts ==="
            try {
                $subAccounts = @(Get-AllSubAccountsRecursive -RootId $RootAccountId)
                Write-Info "Discovered $($subAccounts.Count) sub-accounts"
                foreach ($sub in $subAccounts) {
                    try {
                        Backup-SubAccount -AccountId $sub.id
                    }
                    catch {
                        Write-Fail "Sub-account backup failed for account $($sub.id) ('$($sub.name)'): $_"
                    }
                }
            }
            catch {
                Write-Fail "Failed to enumerate sub-accounts: $_"
            }
        }

        if ($CourseId) {
            Write-Warn "Skipping coverage verification in single-course mode"
            $coverage = [pscustomobject]@{
                CoveragePercent = 100
                CanvasCourses   = 1
                S3Courses       = 1
            }
        }
        else {
            $semester = switch ($today.Month) {
                { $_ -in 1..5 } { 'spring' }
                { $_ -in 6..8 } { 'summer' }
                default { 'fall' }
            }
            $coverage = Test-BackupCoverage -TermIds $termIds -Year $thisYear.ToString() -Semester $semester
        }

        $duration = (Get-Date) - $startTime
        Write-Info ''
        Write-Info '=== SUMMARY ==='
        Write-Info "Duration : $([int]$duration.TotalMinutes) min $($duration.Seconds) sec"
        Write-Info "Content  : written=$cW  skipped=$cS  failed=$cF  dryrun=$cD"
        Write-Info "Gradebook: written=$gW  skipped=$gS  failed=$gF  dryrun=$gD"
        Write-Info "coverage_percent=$($coverage.CoveragePercent)  canvas_courses=$($coverage.CanvasCourses)  s3_courses=$($coverage.S3Courses)"

        if ($WhatIfMode) {
            $runStatus = 'DRYRUN'
            Write-Warn 'BACKUP DRY RUN COMPLETE'
        }
        elseif ($coverage.CoveragePercent -lt 90) {
            $runStatus = 'FAILED'
            Write-Fail 'BACKUP FAILURE'
        }
        else {
            $runStatus = 'COMPLETE'
            Write-Ok 'BACKUP COMPLETE'
        }

        $logs = @(Get-ChildItem $LogDir -Filter '*.log' | Sort-Object LastWriteTime -Descending)
        if ($logs.Count -gt 30) {
            $logs | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        if ($transcriptStarted) {
            Stop-Transcript | Out-Null
        }

        $endTime = Get-Date
        $duration = $endTime - $startTime
        $summaryText = @(
            "run_status=$runStatus"
            "start_time=$($startTime.ToString('s'))"
            "end_time=$($endTime.ToString('s'))"
            "duration_minutes=$([int]$duration.TotalMinutes)"
            "duration_seconds=$($duration.Seconds)"
            "content_written=$cW"
            "content_skipped=$cS"
            "content_failed=$cF"
            "content_dryrun=$cD"
            "gradebook_written=$gW"
            "gradebook_skipped=$gS"
            "gradebook_failed=$gF"
            "gradebook_dryrun=$gD"
            "coverage_percent=$(if ($coverage) { $coverage.CoveragePercent } else { '' })"
            "canvas_courses=$(if ($coverage) { $coverage.CanvasCourses } else { '' })"
            "s3_courses=$(if ($coverage) { $coverage.S3Courses } else { '' })"
            "local_log=$transcriptPath"
            "s3_log_key=$s3LogKey"
            "course_id_filter=$CourseId"
            "max_courses=$MaxCourses"
            "skip_subaccounts=$([bool]$SkipSubAccounts)"
            "skip_content=$([bool]$SkipContent)"
            "skip_gradebooks=$([bool]$SkipGradebooks)"
            "whatif_mode=$([bool]$WhatIfMode)"
        ) -join "`r`n"

        try {
            if (Test-Path $transcriptPath) {
                Publish-FileToS3 -LocalPath $transcriptPath -Key $s3LogKey -StorageClass 'STANDARD'
            }
        }
        catch {
            Write-Warn "Failed to upload transcript log to S3: $_"
        }

        try {
            Publish-TextToS3 -Text $summaryText -Key $s3SummaryKey
        }
        catch {
            Write-Warn "Failed to upload summary file to S3: $_"
        }
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

if (-not $RootAccountId) {
    if ($env:CANVAS_ACCOUNT_ID) {
        $RootAccountId = $env:CANVAS_ACCOUNT_ID
    }
    else {
        $RootAccountId = '1'
    }
}

Main