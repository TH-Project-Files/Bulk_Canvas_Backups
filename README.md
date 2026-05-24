# Canvas LMS Nightly Backup to AWS S3

A robust, self-contained PowerShell script to automate nightly backups of Canvas LMS course content (`.imscc`), Pages (`.json`), Files (`.zip`), and student gradebooks (`.csv`) directly to Amazon S3. Built purely on PowerShell 5.1 and the AWS CLI, it requires no Node.js, NPM, or Python dependencies.

## 🚀 Features & Logic

* **Smart Term Discovery:** Automatically resolves and targets active enrollment terms from the current and prior calendar year.
* **Comprehensive Backup Modalities:** Extracts `.imscc` content packages, HTML-body JSON pages, raw zipped course files, and FERPA-compliant gradebook CSVs.
* **Parallel Processing:** Uses 5 parallel background jobs to handle up to 25 concurrent content exports, vastly speeding up large batch runs.
* **Persistent Sub-account Archiving:** Optionally recurse through your Canvas sub-account tree to ensure older, non-term-bound courses are backed up persistently.
* **Tiered S3 Retention:** Automatically manages storage costs by promoting daily backups to Weekly (retained 5 weeks) and Monthly (retained 6 months) tiers, pruning older files natively via the AWS CLI.
* **Coverage Verification:** Compares the number of courses in Canvas to the number of unique course prefixes in S3, failing the job if coverage drops below 90%.
* **Encrypted Local Config:** Uses Windows DPAPI through CLIXML to protect the Canvas API token at rest.
* **Single-Course Test Mode:** Lets you validate one course quickly without scanning the entire root account.
* **Dry-Run Mode:** Calculates everything, logs the intended actions, and avoids real writes to Canvas export jobs or S3.

---

## 📋 Prerequisites

1. **Windows Environment:** PowerShell 5.1+ (native to Windows Server / Windows 10+).
2. **AWS CLI v2:** Installed and configured in the system PATH.
3. **Canvas API Token:** A Canvas access token with sufficient admin read/export capabilities.
4. **S3 Bucket:** An S3 bucket with appropriate IAM permissions (write/delete/list) and server-side encryption enabled (recommended for FERPA compliance).

---

## 🛠️ Setup & Configuration

### 1. Install AWS CLI v2
Install the AWS CLI v2 on Windows. Verify installation:
```powershell
aws --version
```

### 2. Configure AWS Authentication & Optimization
Ensure the Windows user account running the script is authenticated to AWS, and configure the concurrency settings to support the script's parallel background jobs.

```powershell
aws configure set aws_access_key_id YOUR_KEY
aws configure set aws_secret_access_key YOUR_SECRET
aws configure set default.region us-east-1
aws configure set default.output json

# Required for parallel pushes
aws configure set default.s3.max_concurrent_requests 50
aws configure set default.s3.max_queue_size 10000
```

### 3. Deploy the Script
Create a working directory, typically `C:\Scripts\`, and place `Invoke-CanvasNightlyBackup.ps1` inside it. The script will automatically create `C:\Scripts\Logs\` and `C:\Scripts\Temp\`.

### 4. Generate Encrypted Configuration
To avoid passing tokens in plaintext, generate an encrypted CLIXML config file using Windows DPAPI. **Run this as the same Windows user account that will execute the scheduled task.**

```powershell
$configPath = 'C:\Scripts\canvas-backup-config.clixml'

$config = [pscustomobject]@{
    CanvasBaseUrl   = Read-Host 'Canvas Base URL (e.g., [https://example.instructure.com](https://example.instructure.com))'
    S3Bucket        = Read-Host 'S3 Bucket Name'
    AwsRegion       = Read-Host 'AWS Region (e.g., us-east-1)'
    RootAccountId   = Read-Host 'Root Account ID (Default is 1)'
    CanvasApiToken  = Read-Host 'Canvas API Token' -AsSecureString
}

$config | Export-Clixml -Path $configPath
Write-Host "Encrypted config written to $configPath"
```

### 5. Schedule via Windows Task Scheduler
**Critical Requirement:** Because the `canvas-backup-config.clixml` file is encrypted using Windows DPAPI, the scheduled task **must** execute under the exact same Windows user account that generated the file in Step 4. If a different user (or the SYSTEM account) attempts to run the script, it will fail to decrypt the Canvas API token.

Run the following in an elevated PowerShell prompt to register the nightly job. You will be prompted to enter the password for the executing user account:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NonInteractive -File "C:\Scripts\Invoke-CanvasNightlyBackup.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At '3:00AM'
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 4)
$runAsUser = $env:USERNAME # Ensure this is the same user from Step 4

Register-ScheduledTask -TaskName 'Canvas Nightly Backup' `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -User $runAsUser `
    -Force
```

---

## 📁 Output Locations

Logs are written to `C:\Scripts\Logs\` and temporary processing data to `C:\Scripts\Temp\`.

### S3 Paths
| Data Type | S3 Path Format |
| :--- | :--- |
| **Content** | `canvas-backups/{year}/{semester}/{dept}/{id}/daily/{date}.imscc` |
| **Pages** | `canvas-backups/{year}/{semester}/{dept}/{id}-pages/daily/{date}.json` |
| **Files** | `canvas-backups/{year}/{semester}/{dept}/{id}-files/daily/{date}.zip` |
| **Gradebooks** | `canvas-backups/{year}/{semester}/{dept}/{id}-gradebook/daily/{date}.csv` |
| **Logs** | `canvas-backups/logs/YYYY/MM/backup-YYYY-MM-DD_HHMMSS.log` |

*(Note: Sub-account backups are routed to a `canvas-backups/persistent/{account-slug}/...` prefix).*

---

## 📌 Parameter Reference

| Parameter | Action | Example |
| :--- | :--- | :--- |
| **`-CourseId`** | Targets a single course. | `-CourseId 12345` |
| **`-MaxCourses`** | Limits total courses processed. | `-MaxCourses 5` |
| **`-SkipSubAccounts`** | Bypasses recursive sub-accounts. | `-SkipSubAccounts` |
| **`-SkipContent`** | Skips `.imscc` package backups. | `-SkipContent` |
| **`-SkipPages`** | Skips `.json` HTML page backups. | `-SkipPages` |
| **`-SkipFiles`** | Skips `.zip` raw file backups. | `-SkipFiles` |
| **`-SkipGradebooks`** | Skips `.csv` gradebook backups. | `-SkipGradebooks` |
| **`-WhatIfMode`** | Dry run (no S3/Canvas writes). | `-WhatIfMode` |
| **`-Force`** | Ignores same-day deduplication. | `-Force` |
| **`-ConcurrentJobs`** | Increase subworkers for larger sites, default is 5. | `-ConcurrentJobs 10` |

---

## 🔒 Security & Privacy Notes (FERPA)

* **No Hardcoded Secrets:** AWS relies on standard CLI profiles/roles, and the Canvas token is secured via Windows DPAPI.
* **Ephemeral Data:** Export packages and CSVs containing student data are built in `C:\Scripts\Temp\` and forcefully removed after use.
* **No Public S3 Access:** Ensure your S3 bucket blocks public access.
* **Encrypted Config:** `canvas-backup-config.clixml` should never be committed to source control.

---

## 📝 Appendix: S3 Bucket & IAM Setup Guide

To run this script securely, you need an S3 bucket configured for private access and an IAM user with least-privilege permissions.

### 1. Create the S3 Bucket
1. Log in to the AWS Management Console and navigate to **S3**.
2. Click **Create bucket**.
3. **Bucket name:** Choose a globally unique name (e.g., `school-canvas-backups-2026`).
4. **AWS Region:** Select your preferred region (e.g., `us-east-1`).
5. **Block Public Access settings:** Ensure **Block *all* public access** is checked (Critical for FERPA).
6. **Bucket Versioning:** Leave *Disabled*. (The script natively handles versioning via tiered folders; enabling AWS versioning will result in hidden storage costs when the script automatically prunes older backups).
7. **Default encryption:** Ensure Server-side encryption is enabled (Amazon S3 managed keys - SSE-S3 is sufficient).
8. Click **Create bucket**.

### 2. Create the Least-Privilege IAM Policy
The script requires specific permissions to list objects, upload files, copy objects (for tier promotion), and delete objects (for pruning).

1. Navigate to **IAM** > **Policies** and click **Create policy**.
2. Switch to the **JSON** tab and paste the following policy. Replace `YOUR_BUCKET_NAME` with the bucket name you created in Step 1.

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Effect": "Allow",
            "Action": [
                "s3:ListBucket"
            ],
            "Resource": "arn:aws:s3:::YOUR_BUCKET_NAME"
        },
        {
            "Effect": "Allow",
            "Action": [
                "s3:PutObject",
                "s3:GetObject",
                "s3:DeleteObject"
            ],
            "Resource": "arn:aws:s3:::YOUR_BUCKET_NAME/*"
        }
    ]
}
```
3. Click **Next**, name the policy (e.g., `CanvasBackupS3Access`), and click **Create policy**.

### 3. Create the IAM Service Account & Access Keys
1. Navigate to **IAM** > **Users** and click **Create user**.
2. **User name:** Enter a descriptive name (e.g., `svc-canvas-backup`). Click Next.
3. Under Permissions, select **Attach policies directly**.
4. Search for and select the `CanvasBackupS3Access` policy you just created. Click Next, then **Create user**.
5. Once created, click on the user's name to view their profile.
6. Go to the **Security credentials** tab.
7. Scroll down to **Access keys** and click **Create access key**.
8. Select **Command Line Interface (CLI)**, acknowledge the warning, and proceed.
9. **CRITICAL:** Copy the **Access key ID** and **Secret access key**. This is the only time you can view the Secret Key.
10. Open your Windows Server/Machine and run `aws configure` (as shown in Step 2 of the main Setup Guide) using these newly generated keys.

## License

© 2026 TH-Project-Files. 

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)** License. 

**What this means:**
* **Anyone can use it:** You are free to copy, redistribute, remix, and build upon this framework.
* **Attribute the author:** You must give appropriate credit, provide a link to the license, and indicate if changes were made.
* **No commercial use:** You may not use this material, or derivatives of it, for commercial purposes or monetization.

The software is provided “as is”, without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.

For the full legal terms, please review the [LICENSE.md](LICENSE.md) file included in this repository.
