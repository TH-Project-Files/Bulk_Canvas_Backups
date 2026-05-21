# Canvas LMS Nightly Backup to AWS S3

A robust, self-contained PowerShell script to automate nightly backups of Canvas LMS course content (`.imscc` packages) and student gradebooks (`.csv`) directly to Amazon S3. Built purely on PowerShell 5.1 and the AWS CLI, it requires no Node.js, NPM, or Python dependencies.

## 🚀 Features & Logic

* **Smart Term Discovery:** Automatically resolves and targets active enrollment terms from the current and prior calendar year.
* **Dual Backup Modalities:**
  * **Course Content:** Triggers Canvas backend jobs to generate Common Cartridge (`.imscc`) exports, waits for completion, and archives them.
  * **Gradebooks:** Aggregates assignments, submissions, and enrollment data into a FERPA-compliant CSV matching the native Canvas gradebook export format.
* **Persistent Sub-account Archiving:** Optionally recurse through your Canvas sub-account tree to ensure older, non-term-bound courses are backed up persistently.
* **Tiered S3 Retention:** Automatically manages storage costs by promoting daily backups to Weekly (retained 5 weeks) and Monthly (retained 6 months) tiers, pruning older files natively via the AWS CLI.
* **Coverage Verification:** Compares the number of courses in Canvas to the number of unique course prefixes in S3, failing the job if coverage drops below 90%.
* **Encrypted Local Config:** Uses Windows DPAPI through CLIXML to protect the Canvas API token at rest.
* **Single-Course Test Mode:** Lets you validate one course quickly without scanning the entire root account.
* **Dry-Run Mode:** Calculates everything, logs the intended actions, and avoids real writes to Canvas export jobs or S3.

## 📋 Prerequisites

1. **Windows Environment:** PowerShell 5.1+ (native to Windows Server / Windows 10+).
2. **AWS CLI v2:** Installed and configured.
3. **Canvas API Token:** A Canvas access token with sufficient admin read/export capabilities.
4. **S3 Bucket:** An S3 bucket with appropriate IAM permissions (write/delete/list) and server-side encryption enabled (recommended for FERPA compliance).

## 🛠️ Setup & Configuration

### 1. Install AWS CLI v2

Install the AWS CLI v2 on Windows.

MSI download:

<https://awscli.amazonaws.com/AWSCLIV2.msi>

Verify installation:

```powershell
aws --version
```

Expected output should look similar to:

```text
aws-cli/2.x.x Python/... Windows/...
```

### 2. Configure AWS Authentication

Ensure the Windows user account running the script is authenticated to AWS.

```powershell
aws configure set aws_access_key_id YOUR_KEY
aws configure set aws_secret_access_key YOUR_SECRET
aws configure set default.region us-east-1
aws configure set default.output json
```

Verify access:

```powershell
aws s3 ls
```

> Alternatively, if running on EC2, you can use an attached IAM role instead of static access keys.

### 3. Deploy the Script

Create a working directory, typically `C:\Scripts\`, and place `Invoke-CanvasNightlyBackup.ps1` inside it.

The script will automatically create:

```text
C:\Scripts\Logs\
C:\Scripts\Temp\
```

### 4. Generate Encrypted Configuration

To avoid passing tokens in plaintext, generate an encrypted CLIXML config file using Windows DPAPI.

**Important:** run this as the same Windows user account that will execute the scheduled task.

```powershell
$configPath = 'C:\Scripts\canvas-backup-config.clixml'

$config = [pscustomobject]@{
    CanvasBaseUrl   = Read-Host 'Canvas Base URL (e.g., https://example.instructure.com)'
    S3Bucket        = Read-Host 'S3 Bucket Name'
    AwsRegion       = Read-Host 'AWS Region (e.g., us-east-1)'
    RootAccountId   = Read-Host 'Root Account ID (Default is 1)'
    CanvasApiToken  = Read-Host 'Canvas API Token' -AsSecureString
}

$config | Export-Clixml -Path $configPath
Write-Host "Encrypted config written to $configPath"
```

You can verify the file loads:

```powershell
Import-Clixml 'C:\Scripts\canvas-backup-config.clixml'
```

Expected output example:

```text
CanvasBaseUrl  : https://example.instructure.com
S3Bucket       : example-canvas-backups
AwsRegion      : us-east-1
RootAccountId  : 1
CanvasApiToken : System.Security.SecureString
```

### 5. Schedule via Windows Task Scheduler

Run the following in an elevated PowerShell prompt to register the nightly job:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NonInteractive -File "C:\Scripts\Invoke-CanvasNightlyBackup.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At '3:00AM'
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 4)

Register-ScheduledTask -TaskName 'Canvas Nightly Backup' `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force
```

Run immediately:

```powershell
Start-ScheduledTask -TaskName 'Canvas Nightly Backup'
```

Check last run info:

```powershell
Get-ScheduledTaskInfo -TaskName 'Canvas Nightly Backup'
```

---

## 🧪 Testing & Workflows

The script includes several parameters designed for safe testing without touching production storage or looping through thousands of courses.

### Dry Run (`-WhatIfMode`)

Simulates the full workflow. Resolves terms, lists courses, and reads from Canvas, but **does not**:

* submit Canvas export jobs
* upload files to S3
* copy/promote S3 objects
* delete/prune S3 objects

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -WhatIfMode
```

### Single-Course Dry Run

Target one specific course for quick validation.

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts -WhatIfMode
```

This will:

* fetch only course `12345`
* calculate content and gradebook S3 paths
* skip sub-account recursion
* skip real writes

### Single-Course Real Run

Useful for validating one real course before production rollout.

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts
```

This will:

* back up content if not already done today
* back up gradebook if not already done today
* skip sub-account recursion

### Limit Scope to First N Courses

Useful for controlled testing in normal active-term mode.

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -MaxCourses 5 -SkipSubAccounts -WhatIfMode
```

### Skip Specific Workflows

Skip content exports:

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -SkipContent
```

Skip gradebooks:

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -SkipGradebooks
```

Skip recursive sub-accounts:

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -SkipSubAccounts
```

Example combined test:

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts -SkipGradebooks
```

### Force a Fresh Backup

Bypass same-day deduplication:

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts -Force
```

---

## 📁 Output Locations

### Local Files

Logs are written to:

```text
C:\Scripts\Logs\
```

Temporary downloads are written to:

```text
C:\Scripts\Temp\
```

### S3 Files

Content backups:

```text
canvas-backups/{year}/{semester}/{dept}/{course-identifier}/daily/{date}.imscc
```

Gradebook backups:

```text
canvas-backups/{year}/{semester}/{dept}/{course-identifier}-gradebook/daily/{date}.csv
```

Persistent sub-account backups:

```text
canvas-backups/persistent/{account-slug}/{year}/{semester}/{dept}/{course-identifier}/daily/{date}.imscc
canvas-backups/persistent/{account-slug}/{year}/{semester}/{dept}/{course-identifier}-gradebook/daily/{date}.csv
```

Logs:

```text
canvas-backups/logs/YYYY/MM/backup-YYYY-MM-DD_HHMMSS.log
canvas-backups/logs/YYYY/MM/backup-YYYY-MM-DD_HHMMSS.summary.txt
```

---

## 📌 Parameter Reference

### `-CourseId`

Single-course mode.

Example:

```powershell
-CourseId 12345
```

### `-MaxCourses`

Only process the first N discovered courses in normal mode.

Example:

```powershell
-MaxCourses 5
```

### `-SkipSubAccounts`

Skip recursive sub-account processing.

Example:

```powershell
-SkipSubAccounts
```

### `-SkipContent`

Skip `.imscc` content export backups.

Example:

```powershell
-SkipContent
```

### `-SkipGradebooks`

Skip gradebook CSV backups.

Example:

```powershell
-SkipGradebooks
```

### `-WhatIfMode`

Dry run: no Canvas export creation, no S3 writes, no copies, no deletes.

Example:

```powershell
-WhatIfMode
```

### `-Force`

Ignore same-day dedupe and create fresh exports anyway.

Example:

```powershell
-Force
```

---

## 🔒 Security & Privacy Notes (FERPA)

* **No Hardcoded Secrets:** AWS relies on standard CLI profiles/roles, and the Canvas token is secured via Windows DPAPI.
* **Ephemeral Data:** Export packages and gradebook CSVs containing student data are built in `C:\Scripts\Temp\` and forcefully removed after use.
* **No Public S3 Access:** Ensure your S3 bucket blocks public access.
* **Encrypted Config:** `canvas-backup-config.clixml` should never be committed to source control.

```

---

## 🧯 Troubleshooting

### `401 Unauthorized`

Possible causes:

* bad Canvas token
* wrong Canvas URL
* config file created by different Windows user
* token lacks required access

Test manually:

```powershell
$token = Read-Host 'Paste Canvas token'

Invoke-WebRequest `
  -Uri 'https://example.instructure.com/api/v1/users/self' `
  -Headers @{ Authorization = "Bearer $token" } `
  -UseBasicParsing
```

### `Key not valid for use in specified state`

Cause:

* CLIXML file created by a different Windows account

Fix:

* recreate `canvas-backup-config.clixml` as the same Windows account that runs the script

### Course not found in single-course mode

Example:

```text
CourseId 99999 was not found or is not accessible
```

Fix:

* use a known valid course ID
* verify access manually against `/api/v1/courses/{id}`

### AWS upload failure

Test:

```powershell
aws s3 ls
```

Confirm:

* AWS CLI is installed
* credentials are valid
* bucket exists
* IAM policy allows required S3 actions

---

## ✅ Recommended Validation Sequence

### 1. Confirm AWS works

```powershell
aws --version
aws s3 ls
```

### 2. Create encrypted config

```powershell
$configPath = 'C:\Scripts\canvas-backup-config.clixml'

$config = [pscustomobject]@{
    CanvasBaseUrl   = Read-Host 'Canvas Base URL'
    S3Bucket        = Read-Host 'S3 Bucket'
    AwsRegion       = Read-Host 'AWS Region'
    RootAccountId   = Read-Host 'Root Account ID (Default is 1)'
    CanvasApiToken  = Read-Host 'Canvas API Token' -AsSecureString
}

$config | Export-Clixml -Path $configPath
```

### 3. Single-course dry run

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts -WhatIfMode
```

### 4. Single-course real run

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -CourseId 12345 -SkipSubAccounts
```

### 5. Small limited dry run

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1 -MaxCourses 2 -SkipSubAccounts -WhatIfMode
```

### 6. Full normal run

```powershell
C:\Scripts\Invoke-CanvasNightlyBackup.ps1
```

---
