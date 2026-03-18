# Windows Backup

A comprehensive, production-ready automated backup solution for Windows 11, written in PowerShell. Backs up your Documents folder and VirtualBox VMs to two local drives and Backblaze B2 cloud storage, with full SHA256 integrity verification, DKIM-signed email notifications, and restore verification on every run.

---

## Features

- **3-2-1 backup strategy** — two independent local destinations plus off-site cloud storage
- **Versioned Documents backups** — configurable number of timestamped copies retained on each destination; oldest versions pruned automatically
- **Flat VM backups** — single-copy backup of each VM directory; running VMs are detected and skipped automatically (safe, never touches a live VM)
- **SHA256 integrity verification** — every file copied to a local destination is hashed and compared against its source; any mismatch is flagged immediately
- **Restore firetest** — after every run, a randomly selected file from each local destination is re-hashed against its source to confirm the backup is actually readable and complete
- **Parallel local copy** — both local destinations (D and E) are written simultaneously using PowerShell runspaces, significantly reducing total backup time
- **DKIM-signed email notifications** — backup started, backup completed, and missed-schedule alerts are all sent as RFC 6376-compliant DKIM-signed emails, satisfying strict DMARC (`p=reject; adkim=s`) policies
- **SMTP password in Windows Credential Manager** — the SMTP password is never stored in any script file; it is stored securely via DPAPI and retrieved at runtime
- **RSA-2048 DKIM key in Windows Certificate Store** — private key stored as NonExportable via the CNG Key Storage Provider; never written to disk as a readable file
- **Missed-schedule detection** — if the last successful run was more than a configurable number of hours ago, a separate alert email is sent before the run proceeds
- **Automatic log rotation** — log files older than a configurable number of days are deleted at the end of each run
- **Elapsed time reporting** — every completion log line and email report includes total run duration
- **Compatible with PowerShell 5.1 and PowerShell 7+** — no external modules required

---

## Prerequisites

| Dependency | Purpose | Download |
|---|---|---|
| PowerShell 5.1+ | Included with Windows 10/11 | [PowerShell 7](https://github.com/PowerShell/PowerShell/releases) (optional upgrade) |
| Oracle VirtualBox | VM state detection via `VBoxManage` | [virtualbox.org](https://www.virtualbox.org/) |
| rclone | Upload compressed archives to Backblaze B2 | [rclone.org](https://rclone.org/downloads/) |
| 7-Zip | Compress sources before B2 upload | [7-zip.org](https://www.7-zip.org/) |
| Backblaze B2 account | Cloud storage destination | [backblaze.com](https://www.backblaze.com/b2/) |
| SMTP mailbox | Sending backup notifications | Any provider supporting SMTP AUTH |

---

## Files

| File | Description |
|---|---|
| `WindowsBackup.ps1` | Main backup script — configure this before running |
| `Set-SmtpCredential.ps1` | One-time setup script — stores your SMTP password in Windows Credential Manager |

---

## Quick Start

### 1. Clone or download

```powershell
git clone https://github.com/alan-berger/windows-backup.git
cd windows-backup
```

### 2. Configure the script

Open `WindowsBackup.ps1` in any text editor. All configurable settings are in the **CONFIGURATION BLOCK** at the top of the file (the first ~100 lines). Fill in every variable:

```powershell
# Backup identity — appears in all email subjects and report headers
$BackupName         = "WindowsBackup"   # e.g. "HomeBackup", "AliceBackup"

# Source paths
$SourceDocuments    = "C:\Users\YourUsername\Documents"
$SourceVMs          = "C:\Users\YourUsername\VirtualBox VMs"

# Backup destinations (local drives)
$DestD              = "D:\Backup\WindowsBackup"
$DestE              = "E:\Backup\WindowsBackup"

# Backblaze B2
$RcloneRemoteName   = "b2remote"         # must match your rclone config
$B2BucketName       = "your-bucket-name"

# Email
$EmailFrom          = "backup@yourdomain.com"
$EmailTo            = "you@yourdomain.com"
$SmtpUsername       = "backup@yourdomain.com"
$DkimDomain         = "yourdomain.com"   # must match domain in $EmailFrom
```

See the [Configuration Reference](#configuration-reference) section below for a description of every variable.

### 3. Configure rclone for Backblaze B2

```
rclone config
```

Select **n** (new remote), name it to match `$RcloneRemoteName` (e.g. `b2remote`), choose **Backblaze B2** as the storage type, and enter your B2 Application Key ID and Key. Accept defaults for everything else.

Verify the connection:

```
rclone lsf b2remote:your-bucket-name/
```

### 4. Store your SMTP password

Run the credential setup script **as the same user account** that will run the backup task:

```powershell
.\Set-SmtpCredential.ps1
```

You will be prompted to enter the SMTP password with masked input. The password is stored in Windows Credential Manager under the target name configured in `$SmtpCredentialTarget` and is never written to any file.

> **Gmail users:** use a 16-character [App Password](https://support.google.com/accounts/answer/185833), not your Google account password. App Passwords require 2-Step Verification to be enabled on your account.

### 5. Self-sign the scripts (recommended)

Signing the scripts with a code-signing certificate lets you run them under the `AllSigned` execution policy instead of relying on `-ExecutionPolicy Bypass`. This is the more secure long-term configuration: only scripts that carry a valid, trusted signature will execute.

#### 5a. Create a self-signed code-signing certificate

Run the following in an elevated PowerShell session (Run as Administrator):

```powershell
# Create a code-signing certificate in the current user's personal store
$cert = New-SelfSignedCertificate `
    -Type          CodeSigningCert `
    -Subject       "CN=WindowsBackupCodeSigning" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -HashAlgorithm SHA256 `
    -NotAfter      (Get-Date).AddYears(10)

Write-Host "Certificate thumbprint: $($cert.Thumbprint)"
```

#### 5b. Trust the certificate

For PowerShell's `AllSigned` policy to accept the signature, the signing certificate must be present in both the **Trusted Publishers** and **Trusted Root Certification Authorities** stores. Because this is a self-signed certificate it acts as its own root.

```powershell
# Trust as a publisher (required for AllSigned)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    "TrustedPublisher", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

# Trust the root (suppresses "untrusted root" warnings)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    "Root", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "Certificate trusted."
```

> **Task Scheduler note:** if your scheduled task runs as a different account than your interactive session (e.g. a service account), repeat steps 5a and 5b while logged in as — or impersonating — that account, so the certificate is trusted in that user's stores.

#### 5c. Sign both scripts

The signing commands need to know where your scripts are. Use whichever approach suits you:

**Option A — change to the script directory first, then use relative paths:**

```powershell
cd "C:\Path\To\Your\Scripts"

# Retrieve the certificate (if $cert is no longer in scope, load it by thumbprint)
# $cert = Get-Item "Cert:\CurrentUser\My\<thumbprint from step 5a>"

Set-AuthenticodeSignature -FilePath ".\WindowsBackup.ps1"      -Certificate $cert
Set-AuthenticodeSignature -FilePath ".\Set-SmtpCredential.ps1" -Certificate $cert
```

**Option B — use full absolute paths from anywhere:**

```powershell
Set-AuthenticodeSignature -FilePath "C:\Path\To\Your\Scripts\WindowsBackup.ps1"      -Certificate $cert
Set-AuthenticodeSignature -FilePath "C:\Path\To\Your\Scripts\Set-SmtpCredential.ps1" -Certificate $cert
```

Each command prints a `Status` field. A value of `Valid` confirms the script was signed successfully.

> **Important:** re-sign both scripts whenever you edit them. A signature becomes invalid the moment a file is modified — the script will refuse to run under `AllSigned` until re-signed.

#### 5d. Set the execution policy

```powershell
# Apply to the current user (no elevation required)
Set-ExecutionPolicy -ExecutionPolicy AllSigned -Scope CurrentUser
```

Verify that both scripts are now accepted:

```powershell
Get-AuthenticodeSignature ".\WindowsBackup.ps1"
Get-AuthenticodeSignature ".\Set-SmtpCredential.ps1"
```

Both should show `Status : Valid`.

#### 5e. Update your Task Scheduler action

Once the scripts are signed and the policy is set to `AllSigned`, update the Task Scheduler action arguments to remove `-ExecutionPolicy Bypass`:

```
-NonInteractive -File "C:\Path\To\WindowsBackup.ps1"
```

If you prefer to keep `Bypass` for simplicity (accepting the slightly looser policy), the scripts will continue to work either way.

---

### 6. First run

```powershell
powershell.exe -File ".\WindowsBackup.ps1"
```

On first run the script will:

1. Generate an RSA-2048 DKIM key pair and install it in the Windows Certificate Store
2. Print the DNS TXT record you must add to your DNS provider (see [DKIM DNS Setup](#dkim-dns-setup) below)
3. Send a `[BACKUP STARTED]` email
4. Perform the full backup
5. Send a `[BACKUP OK]` (or `[BACKUP FAILED]`) completion email with duration and integrity results

### 7. Add the DKIM DNS record

The console output and log file from the first run will contain an entry like:

```
Name  : backup._domainkey.yourdomain.com
Type  : TXT
TTL   : Auto
Value : v=DKIM1; k=rsa; p=MIIBIjANBgkqhk...
```

Add this as a DNS TXT record with your DNS provider. Allow up to 5 minutes for propagation, then verify:

```
nslookup -type=TXT backup._domainkey.yourdomain.com 1.1.1.1
```

### 8. Schedule with Task Scheduler

1. Open **Task Scheduler** → **Create Task**
2. **General:** Run whether user is logged on or not; Run with highest privileges
3. **Triggers:** Daily at your preferred time (e.g. 02:00)
4. **Actions:** Start a program
   - Program: `powershell.exe`
   - Arguments: `-NonInteractive -File "C:\Path\To\WindowsBackup.ps1"` (omit `-ExecutionPolicy Bypass` if you completed step 5)
5. **Settings:** Run task as soon as possible after a scheduled start is missed; Do not start a new instance if already running

---

## DKIM DNS Setup

DKIM (DomainKeys Identified Mail) allows receiving mail servers to cryptographically verify that your backup emails were genuinely sent by you. It is required for strict DMARC compliance (`adkim=s; p=reject`).

### What you need

Three DNS records at your domain:

| Record | Type | Purpose |
|---|---|---|
| `yourdomain.com` | TXT (SPF) | Authorises your sending IP: `v=spf1 a:mail.yourdomain.com -all` |
| `backup._domainkey.yourdomain.com` | TXT (DKIM) | Public key — generated and printed on first run |
| `_dmarc.yourdomain.com` | TXT (DMARC) | Policy: `v=DMARC1; p=reject; adkim=s; aspf=s` |

### Shared hosting note

If your SMTP server presents a TLS certificate for a different hostname than the one you connect to (common with shared hosting), set `$SmtpTlsHostname` to the name on the certificate. To discover it:

```powershell
$tcp = [System.Net.Sockets.TcpClient]::new("mail.yourdomain.com", 465)
$ssl = [System.Net.Security.SslStream]::new($tcp.GetStream(), $false, { $true })
$ssl.AuthenticateAsClient("mail.yourdomain.com")
$ssl.RemoteCertificate.GetNameInfo(
    [System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false)
$ssl.Dispose(); $tcp.Dispose()
```

---

## Configuration Reference

All variables are in the **CONFIGURATION BLOCK** at the top of `WindowsBackup.ps1`.

### Backup Identity

| Variable | Default | Description |
|---|---|---|
| `$BackupName` | `WindowsBackup` | Name used in email subjects, report headers, and log banners. Choose something meaningful, e.g. `HomeBackup` or `AliceBackup`. |

### Source Paths

| Variable | Description |
|---|---|
| `$SourceDocuments` | Full path to the folder to back up with versioning |
| `$SourceVMs` | Full path to the VirtualBox VMs folder |

### Local Destinations

| Variable | Description |
|---|---|
| `$DestD` | Backup root on drive D — created automatically if absent |
| `$DestE` | Backup root on drive E — created automatically if absent |

### Backblaze B2

| Variable | Description |
|---|---|
| `$RcloneRemoteName` | rclone remote name, exactly as shown in `rclone config` |
| `$B2BucketName` | Backblaze B2 bucket name |
| `$RclonePath` | `rclone` if on PATH, otherwise full path to `rclone.exe` |

### Compression

| Variable | Description |
|---|---|
| `$SevenZipPath` | Full path to `7z.exe` |
| `$TempCompressDir` | Staging directory for archives before B2 upload; cleared after each upload |

### Versioning

| Variable | Default | Description |
|---|---|---|
| `$DocumentVersionsToRetain` | `10` | How many timestamped Document backup copies to keep per destination |

### Firetest

| Variable | Default | Description |
|---|---|---|
| `$FiretestMinSizeBytes` | `10KB` | Minimum file size eligible for the restore firetest |
| `$FiretestMaxSizeBytes` | `50MB` | Maximum file size eligible for the restore firetest |

### Logging

| Variable | Default | Description |
|---|---|---|
| `$LogDirectory` | `C:\Logs\WindowsBackup` | Where timestamped log files are stored |
| `$LogRetentionDays` | `30` | Log files older than this many days are deleted |

### SMTP

| Variable | Description |
|---|---|
| `$SmtpServer` | Hostname of your SMTP relay |
| `$SmtpPort` | `587` for STARTTLS, `465` for implicit TLS |
| `$SmtpImplicitTls` | `$false` for port 587 (STARTTLS), `$true` for port 465 — must match `$SmtpPort` |
| `$EmailFrom` | Sending address — must match `$DkimDomain` for strict DMARC alignment |
| `$EmailFromName` | Display name shown in email clients |
| `$EmailTo` | Recipient address for all backup notifications |
| `$SmtpUsername` | SMTP authentication username (usually same as `$EmailFrom`) |
| `$SmtpCredentialTarget` | Target name used to retrieve the password from Windows Credential Manager |
| `$SmtpTlsHostname` | Leave blank unless your SMTP server's TLS cert hostname differs from `$SmtpServer` |

### DKIM

| Variable | Default | Description |
|---|---|---|
| `$DkimCertSubject` | `CN=BackupDKIM` | Subject name of the certificate in the Windows Certificate Store |
| `$DkimCertStore` | `CurrentUser` | `CurrentUser` for named user accounts; `LocalMachine` for SYSTEM/service accounts |
| `$DkimSelector` | `backup` | DKIM selector — the DNS record will be `<selector>._domainkey.<domain>` |
| `$DkimDomain` | | Must exactly match the domain in `$EmailFrom` |

### Scheduling

| Variable | Default | Description |
|---|---|---|
| `$MaxHoursBetweenRuns` | `25` | Hours since the last successful run before a missed-schedule alert is sent |

---

## Email Notifications

The script sends three types of email:

| Subject prefix | When sent | Action required |
|---|---|---|
| `[BACKUP STARTED]` | Immediately when the run begins | None — informational |
| `[BACKUP OK]` | Run completed successfully | None — review counts periodically |
| `[BACKUP WARNING]` | A VM was skipped (running) or 7-Zip warned on an inaccessible path | Review which VM was skipped; consider powering it off before the next run |
| `[BACKUP FAILED]` | Integrity mismatch, firetest failure, rclone error, or missing source/destination | Investigate immediately — open the log file for details |
| `[BACKUP ALERT]` | Missed schedule detected | Check Task Scheduler is enabled; confirm the machine was not powered off |

The completion email includes: timestamp, overall status, run duration, version stamp, list of VMs backed up or skipped, integrity check pass/fail counts, and firetest results for both local destinations.

---

## Understanding the Logs

Logs are written to `$LogDirectory` as `Backup_YYYY-MM-DD_HH-MM-SS.log`.

Each line follows this format:

```
[2026-03-18 01:26:37] [INFO] Message text
```

Severity levels:

| Level | Meaning |
|---|---|
| `INFO` | Normal operation |
| `WARNING` | Non-fatal issue — e.g. VM skipped, 7-Zip access warning |
| `ERROR` | Failure affecting backup integrity — hash mismatch, copy failure, rclone error |

---

## Restore Procedures

### From local drive (D or E)

Documents are stored as plain directory trees under timestamped version folders:

```
D:\Backup\WindowsBackup\
  Documents_2026-03-18_01-26-37\   ← newest
  Documents_2026-03-17_01-26-37\
  ...
```

To restore, copy the contents of the version directory you want back to your Documents folder. No special tools required.

VMs are stored flat under:

```
D:\Backup\WindowsBackup\VirtualBox VMs\VMName\
```

Copy the VM directory to your VirtualBox VMs folder, then in VirtualBox: **Machine → Add** → browse to the `.vbox` file.

### From Backblaze B2

```powershell
# List available versions
rclone lsf b2remote:your-bucket/Documents/

# Download a specific version
rclone copy "b2remote:your-bucket/Documents/Documents_2026-03-17_01-26-37.7z" "C:\Restore\"

# Extract
& "C:\Program Files\7-Zip\7z.exe" x "C:\Restore\Documents_2026-03-17_01-26-37.7z" -o"C:\Restore\Documents"
```

---

## Security Notes

- The **SMTP password** is stored in Windows Credential Manager (DPAPI protected, current-user scope). It is never written to any file. Only the Windows account that ran `Set-SmtpCredential.ps1` can read it.
- The **DKIM private key** is stored as a `NonExportable` key in the Windows Certificate Store via the CNG Key Storage Provider. It cannot be extracted via PFX export.
- Both credentials are accessible only to the Windows user account that created them. The Task Scheduler task must run as that same account.
- The script file itself contains no secrets. It is safe to commit to version control.

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| `Credential 'WindowsBackupSmtp' not found` | Setup script not run, or run as a different user | Run `Set-SmtpCredential.ps1` as the same account used by the scheduled task |
| `SMTP: expected 250, received 500` | UTF-8 BOM sent to SMTP server (should not occur in current version) | Ensure you are running the latest version of the script |
| `RemoteCertificateNameMismatch` | SMTP server's TLS cert hostname differs from `$SmtpServer` | Set `$SmtpTlsHostname` to the name on the server's certificate (see diagnostic command in DKIM DNS Setup) |
| `INTEGRITY MISMATCH` | File modified during backup, or failing drive | Schedule backup when machine is idle; run `chkdsk` on the affected drive |
| `7-Zip exited with fatal code 2` | Source path inaccessible or 7-Zip misconfigured | Check `$SourceDocuments` exists and `$SevenZipPath` points to a valid 7z.exe |
| `VBoxManage : NOT FOUND` | VirtualBox not installed or wrong path | Install VirtualBox or update `$VBoxManagePath` |
| `rclone : NOT FOUND` | rclone not installed or not on PATH | Install rclone or set `$RclonePath` to the full path |
| DKIM `permerror` in received headers | DKIM DNS record not yet published or not propagated | Add the DNS TXT record from the first-run output and wait for propagation |
| Firetest `SKIP` | No files in the configured size range exist in Documents | Adjust `$FiretestMinSizeBytes` / `$FiretestMaxSizeBytes` |

---

## License

MIT — see [LICENSE](LICENSE) for details.
