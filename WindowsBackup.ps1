#Requires -Version 5.1

<#
.SYNOPSIS
    Windows Backup Solution — Documents and VirtualBox VMs to local drives and Backblaze B2.
.DESCRIPTION
    Backs up two configurable source locations (Documents folder and VirtualBox VMs) to:
      - Two local destination drives  (raw file copy + SHA256 integrity verification)
      - Backblaze B2 cloud storage    (compressed .7z archive via rclone)

    Features
    ────────
    - Versioned Documents backups: configurable number of timestamped copies retained
    - Flat (single-copy) VM backups: running VMs are automatically detected and skipped
    - SHA256 integrity check on every file copied to local destinations
    - Restore firetest: a random file is re-hashed after each run to confirm recoverability
    - DKIM-signed email notifications: backup started, backup complete, missed-schedule alert
    - SMTP password stored securely in Windows Credential Manager (never in this file)
    - Parallel copy to both local destinations simultaneously (runspace-based)
    - Automatic log rotation and old-version pruning on all three destinations

    EMAIL / DKIM
    ────────────
    Email is sent via a raw SMTP/STARTTLS or implicit-TLS implementation that constructs
    the RFC 5322 message directly and applies a DKIM-Signature header before transmission,
    satisfying strict DMARC (adkim=s) alignment requirements.

    DKIM KEY STORAGE — WINDOWS CERTIFICATE STORE
    ─────────────────────────────────────────────
    The RSA-2048 DKIM private key is stored as a NonExportable key in the Windows
    Certificate Store (Cert:\CurrentUser\My by default).  The key material never appears
    on disk as a readable file; it is protected by the CNG key storage provider and DPAPI.

    On first run, the script generates the RSA key pair, installs the self-signed
    certificate into the configured store, and prints the exact DNS TXT record that must
    be published before DMARC-signed mail will be accepted.

    IDENTIFYING THE CERTIFICATE
    ───────────────────────────
    The certificate is found by its Subject name ($DkimCertSubject, default "CN=BackupDKIM").
    To inspect or remove it:
      Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=BackupDKIM" }

    TASK SCHEDULER / CERT STORE CHOICE
    ────────────────────────────────────
    $DkimCertStore = "CurrentUser"  — use when the task runs as a named user account
                                      (the recommended default).
    $DkimCertStore = "LocalMachine" — use when running as SYSTEM or a managed service
                                      account.  You must also grant that account read
                                      access to the private key:
                                        certlm.msc → right-click cert →
                                        All Tasks → Manage Private Keys

    SMTP PASSWORD
    ─────────────
    The SMTP password is never stored in this script file.  Run Set-SmtpCredential.ps1
    once to store it securely in Windows Credential Manager before the first backup run.

.NOTES
    Compatible with PowerShell 5.1 and PowerShell 7+
    Version: 3.0
    Repository: https://github.com/YOUR-USERNAME/YOUR-REPO
#>

Set-StrictMode -Version Latest

###############################################################################
#region ======================================================================
#                        CONFIGURATION BLOCK
#   Edit EVERY variable in this section before your first run.
#   Do NOT modify anything outside this region unless you know what you are doing.
# =============================================================================

# ---------------------------------------------------------------------------
# Backup Identity
# This name appears in email subjects, report headers, and log banners.
# Choose something meaningful, e.g. "HomeBackup", "AliceBackup", "OfficeBackup".
# Avoid spaces if you also use this as part of a file or directory name.
# ---------------------------------------------------------------------------
$BackupName         = "WindowsBackup"

# ---------------------------------------------------------------------------
# Source Paths
# Replace with the actual paths you want to back up.
# ---------------------------------------------------------------------------
$SourceDocuments    = "C:\Users\YourUsername\Documents"
$SourceVMs          = "C:\Users\YourUsername\VirtualBox VMs"

# ---------------------------------------------------------------------------
# Local Destination Paths
# These directories are created automatically if they do not exist.
# The parent drive (D:, E:) must be present and accessible.
# ---------------------------------------------------------------------------
$DestD              = "D:\Backup\WindowsBackup"
$DestE              = "E:\Backup\WindowsBackup"

# ---------------------------------------------------------------------------
# Backblaze B2 / rclone Settings
# $RcloneRemoteName must match the remote name you created with `rclone config`.
# ---------------------------------------------------------------------------
$RcloneRemoteName   = "b2remote"
$B2BucketName       = "your-b2-bucket-name"
$RclonePath         = "rclone"          # full path if rclone is not on your PATH

# ---------------------------------------------------------------------------
# Compression Settings (7-Zip)
# $TempCompressDir needs enough free space for one compressed Documents archive
# and one VM archive simultaneously.
# ---------------------------------------------------------------------------
$SevenZipPath       = "C:\Program Files\7-Zip\7z.exe"
$TempCompressDir    = "C:\Temp\BackupCompress"

# ---------------------------------------------------------------------------
# VirtualBox
# Update the path if VirtualBox is installed to a non-default location.
# ---------------------------------------------------------------------------
$VBoxManagePath     = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

# ---------------------------------------------------------------------------
# Versioning (Documents only)
# How many timestamped copies to retain on each destination (D, E, and B2).
# ---------------------------------------------------------------------------
$DocumentVersionsToRetain = 10

# ---------------------------------------------------------------------------
# Firetest File Size Thresholds
# The restore firetest selects a random file within this size range.
# Widen the range if your Documents folder has few files in the 10 KB–50 MB window.
# ---------------------------------------------------------------------------
$FiretestMinSizeBytes = 10KB        # files smaller than this are excluded
$FiretestMaxSizeBytes = 50MB        # files larger than this are excluded

# ---------------------------------------------------------------------------
# Logging
# Log files older than $LogRetentionDays are deleted at the end of each run.
# ---------------------------------------------------------------------------
$LogDirectory       = "C:\Logs\WindowsBackup"
$LogRetentionDays   = 30

# ---------------------------------------------------------------------------
# SMTP Settings
# Configure these to match your mail provider.
# ---------------------------------------------------------------------------
$SmtpServer         = "smtp.example.com"
$SmtpPort           = 587           # 587 = STARTTLS (explicit TLS)  |  465 = implicit TLS

# TLS mode — must match $SmtpPort:
#   $false  ->  STARTTLS  (port 587): connect plain, issue STARTTLS command, then upgrade.
#   $true   ->  Implicit TLS (port 465): TLS handshake fires immediately on TCP connect,
#               before any SMTP conversation begins.  Set $true when using port 465.
$SmtpImplicitTls    = $false

$EmailFrom          = "backup@yourdomain.com"
$EmailFromName      = "Windows Backup"
$EmailTo            = "you@yourdomain.com"
$SmtpUsername       = "backup@yourdomain.com"

# The target name used to look up the SMTP password in Windows Credential Manager.
# Run Set-SmtpCredential.ps1 once to store the password securely.
# The password is NEVER stored in this script file.
$SmtpCredentialTarget = "WindowsBackupSmtp"

# TLS certificate hostname — the CN/SAN on the SMTP server's TLS certificate.
# Leave blank ("") when the certificate hostname matches $SmtpServer exactly (most common).
# Set explicitly when your provider's cert is issued to a different hostname, for example
# when connecting to a shared-hosting SMTP server that presents its own server name:
#   $SmtpServer      = "mail.yourdomain.com"   (the hostname you connect to)
#   $SmtpTlsHostname = "mail.provider.com"     (the name on the server's certificate)
$SmtpTlsHostname    = ""

# ---------------------------------------------------------------------------
# DKIM Signing Settings
# DKIM allows receiving mail servers to verify your email was genuinely sent
# by you, and is required for strict DMARC compliance.
# See README.md for full DNS setup instructions.
# ---------------------------------------------------------------------------

# Subject name of the self-signed certificate created in the Windows Certificate Store.
# The certificate is generated automatically on the first run.
$DkimCertSubject    = "CN=BackupDKIM"

# Which certificate store to use.
#   "CurrentUser"  — recommended when Task Scheduler runs as a named user account.
#   "LocalMachine" — use when running as SYSTEM or a managed service account.
#                    Requires manual read ACL grant to the service account:
#                    certlm.msc -> right-click cert -> All Tasks -> Manage Private Keys
$DkimCertStore      = "CurrentUser"

# DKIM selector — determines the DNS record name:  <selector>._domainkey.<domain>
# Example with selector "backup":  backup._domainkey.yourdomain.com
$DkimSelector       = "backup"

# Must exactly match the domain portion of $EmailFrom.
# With adkim=s (strict DMARC alignment) any mismatch causes delivery failure.
$DkimDomain         = "yourdomain.com"

# ---------------------------------------------------------------------------
# Missed Schedule Detection
# If the most recent successful run was more than this many hours ago when the
# script starts, a separate missed-schedule alert email is sent.
# For a daily schedule, 25 allows a one-hour tolerance window.
# ---------------------------------------------------------------------------
$MaxHoursBetweenRuns = 25

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                          SCRIPT-WIDE STATE
# =============================================================================

$script:RunTimestamp    = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:RunStartTime    = Get-Date      # recorded here; used for elapsed-time calculation
$script:LogFile         = $null
$script:OverallStatus   = "SUCCESS"
$script:SmtpPassword    = $null         # populated at runtime from Windows Credential Manager

$script:IntegrityPassed = 0
$script:IntegrityFailed = 0

$script:SkippedVMs      = [System.Collections.Generic.List[string]]::new()
$script:BackedUpVMs     = [System.Collections.Generic.List[string]]::new()
$script:FiretestResults = [System.Collections.Generic.List[string]]::new()
$script:ErrorsList      = [System.Collections.Generic.List[string]]::new()

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                            LOGGING
# =============================================================================

function Initialize-Log {
    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory "Backup_$($script:RunTimestamp).log"
    Write-Log "INFO" "================================================================"
    Write-Log "INFO" "  Windows Backup v3.0 -- Run started: $($script:RunTimestamp)"
    Write-Log "INFO" "================================================================"
}

function Write-Log {
    param(
        [ValidateSet("INFO","WARNING","ERROR")]
        [string]$Level,
        [string]$Message
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8
    }
    switch ($Level) {
        "INFO"    { Write-Host $entry -ForegroundColor Cyan   }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $entry -ForegroundColor Red    }
    }
}

function Remove-OldLogs {
    Write-Log "INFO" "Purging log files older than $LogRetentionDays days..."
    $cutoff = (Get-Date).AddDays(-$LogRetentionDays)
    try {
        Get-ChildItem -Path $LogDirectory -Filter "Backup_*.log" -ErrorAction Stop |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            ForEach-Object {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                Write-Log "INFO" "Purged old log: $($_.Name)"
            }
    } catch {
        Write-Log "WARNING" "Log purge encountered an error: $_"
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                   WINDOWS CREDENTIAL MANAGER  (SMTP password)
#
#  The SMTP password is stored as a Generic credential in the Windows
#  Credential Manager under the target name $SmtpCredentialTarget.
#  It is never written to this script file.
#
#  To store or update the password, run:
#    Set-SmtpCredential.ps1
#
#  The C# helper class uses direct Win32 P/Invoke (advapi32.dll CredRead).
#  This works on PowerShell 5.1 (.NET Framework 4.x) and PowerShell 7+
#  (.NET 5+) without any external modules.
#
#  Access: the credential is readable only by the Windows account that stored
#  it (CurrentUser scope / DPAPI protection).  The Task Scheduler task must
#  run as the same user account that ran Set-SmtpCredential.ps1.
# =============================================================================

function Initialize-CredentialManager {
    <#
    .SYNOPSIS  Compiles the CredentialManager P/Invoke helper type into the session.
               Uses -TypeDefinition so that C# `using` directives are at file scope,
               which is required — they cannot appear inside a class body (-MemberDefinition).
               Guarded against duplicate compilations within the same process.
    #>
    if (([System.Management.Automation.PSTypeName]'WindowsBackup.CredentialManager').Type) {
        return   # type already loaded in this session
    }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsBackup {
    public class CredentialManager {

        // CREDENTIAL layout must match the Windows CREDENTIAL struct exactly.
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CREDENTIAL {
            public uint   Flags;
            public uint   Type;
            public string TargetName;
            public string Comment;
            public long   LastWritten;
            public uint   CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint   Persist;
            public uint   AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credPtr);

        [DllImport("advapi32.dll")]
        private static extern void CredFree(IntPtr cred);

        // Reads a Generic (type=1) credential from Windows Credential Manager.
        // Returns the password as a plain string.
        public static string ReadPassword(string target) {
            IntPtr ptr = IntPtr.Zero;
            if (!CredRead(target, 1, 0, out ptr))
                throw new InvalidOperationException(
                    "Credential '" + target + "' not found in Windows Credential Manager. " +
                    "Run Set-SmtpCredential.ps1 to store it.");
            try {
                var cred = (CREDENTIAL)Marshal.PtrToStructure(ptr, typeof(CREDENTIAL));
                if (cred.CredentialBlobSize == 0) return string.Empty;
                var bytes = new byte[cred.CredentialBlobSize];
                Marshal.Copy(cred.CredentialBlob, bytes, 0, (int)cred.CredentialBlobSize);
                return Encoding.Unicode.GetString(bytes);
            } finally {
                CredFree(ptr);
            }
        }
    }
}
'@ -ErrorAction Stop
}

function Get-SmtpPasswordFromCredentialManager {
    <#
    .SYNOPSIS  Retrieves the SMTP password from Windows Credential Manager and
               stores it in $script:SmtpPassword for use by Send-DkimSignedEmail.
               Throws with a clear message if the credential does not exist.
    #>
    Write-Log "INFO" "Loading SMTP credentials from Windows Credential Manager..."
    try {
        Initialize-CredentialManager
        $script:SmtpPassword = [WindowsBackup.CredentialManager]::ReadPassword($SmtpCredentialTarget)
        Write-Log "INFO" "  SMTP credential loaded: target='$SmtpCredentialTarget'"
    } catch {
        throw "Cannot load SMTP credential '$SmtpCredentialTarget': $_"
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                         ELAPSED TIME HELPER
# =============================================================================

function Format-Elapsed {
    <#
    .SYNOPSIS  Returns a human-readable elapsed-time string (e.g. "1h 23m 45s").
    #>
    param([System.TimeSpan]$Elapsed)
    $h = [int]$Elapsed.TotalHours
    $m = $Elapsed.Minutes
    $s = $Elapsed.Seconds
    if ($h -ge 1) { return "${h}h ${m}m ${s}s" }
    if ($m -ge 1) { return "${m}m ${s}s" }
    return "${s}s"
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                   BACKUP STARTED EMAIL
# =============================================================================

function Send-BackupStartedEmail {
    <#
    .SYNOPSIS  Sends a brief notification that a backup run has begun.
               Sent before any backup work so you know a run is in progress
               even if the script later fails before sending the final report.
    #>
    $subject = "[BACKUP STARTED] $BackupName -- $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $sep     = "=" * 50
    $body = @"
$sep
  $BackupName -- BACKUP STARTED
$sep
  Started    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Machine    : $env:COMPUTERNAME
  Log File   : $($script:LogFile)

SOURCES
  Documents  : $SourceDocuments
  VMs        : $SourceVMs

DESTINATIONS
  Local D    : $DestD
  Local E    : $DestE
  Cloud      : ${RcloneRemoteName}:${B2BucketName}

The backup completion report will follow when the run finishes.
"@
    Send-DkimSignedEmail -Subject $subject -Body $body
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                       MISSED SCHEDULE DETECTION
# =============================================================================

function Get-LastSuccessfulRunTime {
    $logs = Get-ChildItem -Path $LogDirectory -Filter "Backup_*.log" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            Select-Object -Skip 1

    foreach ($log in $logs) {
        $content = Get-Content $log.FullName -Raw -ErrorAction SilentlyContinue
        if ($content -match "Backup run completed with overall status: SUCCESS") {
            return $log.LastWriteTime
        }
    }
    return $null
}

function Test-MissedSchedule {
    Write-Log "INFO" "--- Missed-schedule check ---"
    $lastRun = Get-LastSuccessfulRunTime

    if ($null -eq $lastRun) {
        Write-Log "INFO" "No previous successful run found -- skipping missed-schedule check."
        return
    }

    $elapsed        = ((Get-Date) - $lastRun).TotalHours
    $elapsedRounded = [math]::Round($elapsed, 1)

    if ($elapsed -gt $MaxHoursBetweenRuns) {
        $msg = "Missed schedule: last successful run was $elapsedRounded hours ago (threshold: $MaxHoursBetweenRuns h)."
        Write-Log "WARNING" $msg
        Send-MissedScheduleAlert -LastRun $lastRun -ElapsedHours $elapsedRounded
    } else {
        Write-Log "INFO" "Schedule OK -- last successful run: $lastRun ($elapsedRounded hours ago)."
    }
}

function Send-MissedScheduleAlert {
    param([datetime]$LastRun, [double]$ElapsedHours)
    $subject = "[BACKUP ALERT] $BackupName -- Missed Schedule -- $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $sep     = "=" * 50
    $body = @"
$sep
  $BackupName -- MISSED SCHEDULE ALERT
$sep
The backup script has detected that the last successful backup
was $ElapsedHours hours ago, which exceeds the configured threshold of
$MaxHoursBetweenRuns hours.

Last successful run : $LastRun
Threshold           : $MaxHoursBetweenRuns hours
Detected at         : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Please verify that the Windows Task Scheduler task is active and that the
machine was not powered off or disconnected during the scheduled window.
Check the log directory for details:
  $LogDirectory
"@
    Send-DkimSignedEmail -Subject $subject -Body $body
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                        DEPENDENCY CHECKS
# =============================================================================

function Assert-Dependencies {
    Write-Log "INFO" "--- Dependency check ---"
    $allOk = $true

    if (Test-Path $VBoxManagePath) {
        Write-Log "INFO" "VBoxManage   : OK  ($VBoxManagePath)"
    } else {
        Write-Log "ERROR" "VBoxManage   : NOT FOUND at $VBoxManagePath"
        $allOk = $false
    }

    $rcloneResolved = Get-Command $RclonePath -ErrorAction SilentlyContinue
    if ($rcloneResolved) {
        Write-Log "INFO" "rclone       : OK  ($($rcloneResolved.Source))"
    } else {
        Write-Log "ERROR" "rclone       : NOT FOUND. Searched: $RclonePath"
        $allOk = $false
    }

    if (Test-Path $SevenZipPath) {
        Write-Log "INFO" "7-Zip        : OK  ($SevenZipPath)"
    } else {
        Write-Log "ERROR" "7-Zip        : NOT FOUND at $SevenZipPath"
        $allOk = $false
    }

    if (-not $allOk) {
        $script:OverallStatus = "FAILURE"
        throw "One or more required dependencies are missing. See log for details."
    }

    Write-Log "INFO" "All dependencies satisfied."
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                       SHA256 INTEGRITY CHECKS
# =============================================================================

function Get-FileHash256 {
    <#
    .SYNOPSIS  Returns the SHA256 hash of a file.
    .NOTES     Uses -LiteralPath so that file names containing PowerShell wildcard
               characters ([ ] * ?) are handled correctly and are never glob-expanded.
    #>
    param([string]$FilePath)
    return (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Test-FileCopyIntegrity {
    param([string]$SourceFile, [string]$DestFile)
    try {
        $srcHash = Get-FileHash256 $SourceFile
        $dstHash = Get-FileHash256 $DestFile

        if ($srcHash -eq $dstHash) {
            $script:IntegrityPassed++
            Write-Log "INFO" "  INTEGRITY OK : $(Split-Path $DestFile -Leaf)  [$srcHash]"
            return $true
        } else {
            $script:IntegrityFailed++
            $msg = "INTEGRITY MISMATCH | Source: $SourceFile | Dest: $DestFile | Expected: $srcHash | Got: $dstHash"
            Write-Log "ERROR" $msg
            $script:ErrorsList.Add($msg)
            Set-OverallStatus "FAILURE"
            return $false
        }
    } catch {
        $script:IntegrityFailed++
        $msg = "INTEGRITY CHECK ERROR | Dest: $DestFile | Error: $_"
        Write-Log "ERROR" $msg
        $script:ErrorsList.Add($msg)
        Set-OverallStatus "FAILURE"
        return $false
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                      STATUS HELPER
# =============================================================================

function Set-OverallStatus {
    param([ValidateSet("SUCCESS","WARNING","FAILURE")][string]$NewStatus)
    $order = @{ "SUCCESS" = 0; "WARNING" = 1; "FAILURE" = 2 }
    if ($order[$NewStatus] -gt $order[$script:OverallStatus]) {
        $script:OverallStatus = $NewStatus
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                     DIRECTORY COPY WITH INTEGRITY
# =============================================================================

function Copy-DirectoryWithIntegrity {
    param([string]$SourceDir, [string]$DestDir)
    Write-Log "INFO" "  Copying: $SourceDir"
    Write-Log "INFO" "       To: $DestDir"

    if (-not (Test-Path -LiteralPath $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }

    # -LiteralPath on the base dir so that a source path containing brackets
    # is not treated as a wildcard pattern by Get-ChildItem.
    $allFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -File -ErrorAction SilentlyContinue

    if ($null -eq $allFiles -or @($allFiles).Count -eq 0) {
        Write-Log "WARNING" "  No files found in: $SourceDir"
        return
    }

    foreach ($file in $allFiles) {
        $relPath     = $file.FullName.Substring($SourceDir.TrimEnd('\').Length).TrimStart('\')
        $destFile    = Join-Path $DestDir $relPath
        $destFileDir = Split-Path $destFile -Parent

        # -LiteralPath so that directory names with brackets are not glob-expanded.
        if (-not (Test-Path -LiteralPath $destFileDir)) {
            New-Item -ItemType Directory -Path $destFileDir -Force | Out-Null
        }

        try {
            # -LiteralPath on the source so files named e.g. "Foo [SN].wav" are
            # copied correctly rather than triggering a wildcard expansion error.
            Copy-Item -LiteralPath $file.FullName -Destination $destFile -Force -ErrorAction Stop
            Write-Log "INFO" "  Copied: $relPath"
            Test-FileCopyIntegrity -SourceFile $file.FullName -DestFile $destFile | Out-Null
        } catch {
            $script:IntegrityFailed++
            $msg = "COPY FAILED | Source: $($file.FullName) | Dest: $destFile | Error: $_"
            Write-Log "ERROR" $msg
            $script:ErrorsList.Add($msg)
            Set-OverallStatus "FAILURE"
        }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                    PARALLEL LOCAL COPY  (D and E simultaneously)
#
#  Design
#  ──────
#  Each local destination (D and E) reads from the same source and writes to
#  an independent drive.  There is no shared state between the two operations,
#  making them safe to run concurrently using PowerShell runspaces.
#
#  $script:ParallelCopyWorker is a self-contained [scriptblock] that performs
#  the full copy + SHA256 integrity loop for a single destination.  It carries
#  no references to script-scope variables; all inputs are parameters and all
#  outputs are returned as a PSCustomObject.  This avoids thread-safety issues
#  entirely — each runspace owns its own data.
#
#  Invoke-ParallelLocalCopy
#  ────────────────────────
#  1. Opens a RunspacePool(1, 2) — max two concurrent runspaces.
#  2. Launches one PowerShell instance per destination (D and E) and starts
#     them both with BeginInvoke().
#  3. Calls EndInvoke() on each handle to collect results.
#  4. Flushes the worker's buffered log entries to the main log file and
#     console, then merges counters and errors into script-scope state.
#
#  B2 always runs sequentially after both local jobs complete.  This avoids
#  source-drive read contention with the local copies, and the upload is
#  network-bound rather than disk-bound anyway.
#
#  Compatibility: PowerShell 5.1 and PowerShell 7+.
#  RunspaceFactory and PowerShell::Create() are available in both.
#
# =============================================================================

$script:ParallelCopyWorker = [scriptblock] {
    <#
    .SYNOPSIS  Self-contained copy + SHA256 integrity worker for one destination.
               Designed to run inside a runspace — no script-scope dependencies.
    .PARAMETER SourceDir  Absolute path of the directory to copy from.
    .PARAMETER DestDir    Absolute path of the directory to copy into.
    .OUTPUTS   PSCustomObject with:
               .Passed     [int]    Files where source and dest hashes matched.
               .Failed     [int]    Files where the hash was missing or mismatched.
               .Errors     [List[string]]  Error messages to merge into ErrorsList.
               .LogEntries [List[string]]  Timestamped log lines to flush to main log.
    #>
    param(
        [string]$SourceDir,
        [string]$DestDir
    )

    $result = [PSCustomObject]@{
        Passed     = 0
        Failed     = 0
        Errors     = [System.Collections.Generic.List[string]]::new()
        LogEntries = [System.Collections.Generic.List[string]]::new()
    }

    # Local helpers — cannot call main-thread functions from inside a runspace.
    function Worker-Log {
        param([string]$Level, [string]$Msg)
        $result.LogEntries.Add("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Msg")
    }

    function Worker-Hash {
        param([string]$Path)
        # -LiteralPath prevents wildcard expansion on filenames with [ ] * ?
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    }

    Worker-Log "INFO" "  [PARALLEL] Copying: $SourceDir"
    Worker-Log "INFO" "  [PARALLEL]       To: $DestDir"

    if (-not (Test-Path -LiteralPath $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }

    $allFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -File -ErrorAction SilentlyContinue

    if ($null -eq $allFiles -or @($allFiles).Count -eq 0) {
        Worker-Log "WARNING" "  [PARALLEL] No files found in: $SourceDir"
        return $result
    }

    foreach ($file in $allFiles) {
        $relPath     = $file.FullName.Substring($SourceDir.TrimEnd('\').Length).TrimStart('\')
        $destFile    = Join-Path $DestDir $relPath
        $destFileDir = Split-Path $destFile -Parent

        if (-not (Test-Path -LiteralPath $destFileDir)) {
            New-Item -ItemType Directory -Path $destFileDir -Force | Out-Null
        }

        try {
            Copy-Item -LiteralPath $file.FullName -Destination $destFile -Force -ErrorAction Stop
            Worker-Log "INFO" "  Copied: $relPath"

            # Integrity check
            try {
                $srcHash = Worker-Hash $file.FullName
                $dstHash = Worker-Hash $destFile

                if ($srcHash -eq $dstHash) {
                    $result.Passed++
                    Worker-Log "INFO" "  INTEGRITY OK : $(Split-Path $destFile -Leaf)  [$srcHash]"
                } else {
                    $result.Failed++
                    $msg = "INTEGRITY MISMATCH | Source: $($file.FullName) | Dest: $destFile | Expected: $srcHash | Got: $dstHash"
                    Worker-Log "ERROR" $msg
                    $result.Errors.Add($msg)
                }
            } catch {
                $result.Failed++
                $msg = "INTEGRITY CHECK ERROR | Dest: $destFile | Error: $_"
                Worker-Log "ERROR" $msg
                $result.Errors.Add($msg)
            }

        } catch {
            $result.Failed++
            $msg = "COPY FAILED | Source: $($file.FullName) | Dest: $destFile | Error: $_"
            Worker-Log "ERROR" $msg
            $result.Errors.Add($msg)
        }
    }

    return $result
}   # end $script:ParallelCopyWorker


function Flush-WorkerResult {
    <#
    .SYNOPSIS  Writes a parallel worker's buffered log entries to the main log
               file and console, then merges its counters and errors into the
               script-scope state.
    .PARAMETER WorkerResult  The PSCustomObject returned by the parallel worker.
    .PARAMETER Destination   Label string used in the section header (e.g. "D:\Backup\WindowsBackup").
    #>
    param(
        [PSCustomObject]$WorkerResult,
        [string]$Destination
    )

    Write-Log "INFO" "  -- Parallel copy results for: $Destination --"

    # Replay buffered log entries — write directly to file and console
    # (Write-Log cannot be called here because it also writes to the file, and
    # we want the console colour to reflect each entry's level).
    foreach ($entry in $WorkerResult.LogEntries) {
        Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8
        $level = if ($entry -match '\[(INFO|WARNING|ERROR)\]') { $Matches[1] } else { 'INFO' }
        $colour = switch ($level) {
            'INFO'    { 'Cyan'   }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red'    }
        }
        Write-Host $entry -ForegroundColor $colour
    }

    # Merge counters
    $script:IntegrityPassed += $WorkerResult.Passed
    $script:IntegrityFailed += $WorkerResult.Failed

    # Merge errors
    foreach ($e in $WorkerResult.Errors) {
        $script:ErrorsList.Add($e)
    }

    if ($WorkerResult.Failed -gt 0) {
        Set-OverallStatus "FAILURE"
    }
}


function Invoke-ParallelLocalCopy {
    <#
    .SYNOPSIS  Copies $SourceDir to both $DestDirD and $DestDirE simultaneously
               using two runspaces, then merges results back into script-scope state.
    .PARAMETER SourceDir  The directory to copy from.
    .PARAMETER DestDirD   Full destination path on drive D.
    .PARAMETER DestDirE   Full destination path on drive E.
    .PARAMETER Label      Human-readable label for log messages (e.g. "Documents" or "VM 'VMK'").
    #>
    param(
        [string]$SourceDir,
        [string]$DestDirD,
        [string]$DestDirE,
        [string]$Label
    )

    Write-Log "INFO" "  Starting parallel copy of $Label to D: and E: simultaneously..."

    # RunspacePool(min=1, max=2) — one runspace per destination, run concurrently.
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 2)
    $pool.Open()

    # Launch both jobs
    $jobs = foreach ($destDir in @($DestDirD, $DestDirE)) {
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($script:ParallelCopyWorker)
        [void]$ps.AddParameter('SourceDir', $SourceDir)
        [void]$ps.AddParameter('DestDir',   $destDir)
        [PSCustomObject]@{
            PS     = $ps
            Handle = $ps.BeginInvoke()
            Dest   = $destDir
        }
    }

    # Collect results in submission order (D first, then E)
    foreach ($job in $jobs) {
        try {
            # EndInvoke blocks until this job finishes; the other job may still be running.
            $returned = $job.PS.EndInvoke($job.Handle)

            # EndInvoke returns a PSDataCollection — the worker's return value is item [0].
            $workerResult = $returned[0]

            # Surface any unhandled runspace errors
            if ($job.PS.HadErrors) {
                foreach ($rsErr in $job.PS.Streams.Error) {
                    $msg = "RUNSPACE ERROR [$($job.Dest)] : $rsErr"
                    Write-Log "ERROR" $msg
                    $script:ErrorsList.Add($msg)
                    Set-OverallStatus "FAILURE"
                }
            }

            Flush-WorkerResult -WorkerResult $workerResult -Destination $job.Dest

        } catch {
            $msg = "PARALLEL COPY FAILED [$($job.Dest)] : $_"
            Write-Log "ERROR" $msg
            $script:ErrorsList.Add($msg)
            Set-OverallStatus "FAILURE"
        } finally {
            $job.PS.Dispose()
        }
    }

    $pool.Close()
    $pool.Dispose()

    Write-Log "INFO" "  Parallel copy of $Label to D: and E: complete."
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                          VERSIONING
# =============================================================================

function Get-VersionedPath {
    param([string]$BaseDir, [string]$SourceName)
    return Join-Path $BaseDir "${SourceName}_$($script:RunTimestamp)"
}

function Remove-OldVersions {
    param([string]$BaseDir, [string]$SourceName, [int]$KeepCount)
    Write-Log "INFO" "  Pruning old versions of '$SourceName' in $BaseDir (retain: $KeepCount)"

    $versions = Get-ChildItem -Path $BaseDir -Directory -Filter "${SourceName}_????-??-??_??-??-??" -ErrorAction SilentlyContinue |
                Sort-Object Name -Descending

    if ($versions.Count -le $KeepCount) {
        Write-Log "INFO" "  No pruning needed ($($versions.Count) version(s) present)."
        return
    }

    $toDelete = $versions | Select-Object -Skip $KeepCount
    foreach ($dir in $toDelete) {
        try {
            Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
            Write-Log "INFO" "  Pruned: $($dir.Name)"
        } catch {
            Write-Log "WARNING" "  Failed to prune $($dir.FullName): $_"
        }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                     B2 VERSION PRUNING
# =============================================================================

function Remove-OldB2Versions {
    param([string]$RemotePath, [string]$FilePrefix, [int]$KeepCount)
    Write-Log "INFO" "  Pruning old B2 versions in ${RemotePath} (retain: $KeepCount)"

    try {
        $lsOutput = & $RclonePath lsf $RemotePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "WARNING" "  rclone lsf failed: $lsOutput"
            return
        }

        # @() wraps the entire pipeline result, preventing PowerShell 5.1 from
        # unwrapping a single-element result into a bare scalar — which would cause
        # the subsequent .Count call to throw "property not found" on a string.
        $files = @(@($lsOutput) | Where-Object { $_ -like "${FilePrefix}*" } | Sort-Object -Descending)

        if ($files.Count -le $KeepCount) {
            Write-Log "INFO" "  No B2 pruning needed ($($files.Count) file(s) present)."
            return
        }

        # @() for the same scalar-unwrap reason as above.
        $toDelete = @($files | Select-Object -Skip $KeepCount)
        foreach ($f in $toDelete) {
            $target = "${RemotePath}$($f.TrimEnd('/'))"
            Write-Log "INFO" "  Deleting old B2 version: $f"
            $delOut = & $RclonePath deletefile $target 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Log "WARNING" "  rclone deletefile failed for $f : $delOut"
            } else {
                Write-Log "INFO" "  Deleted B2 object: $f"
            }
        }
    } catch {
        Write-Log "WARNING" "  B2 version pruning error: $_"
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                 VIRTUALBOX VM STATE CHECK
# =============================================================================

function Get-RunningVMNames {
    Write-Log "INFO" "Querying VirtualBox for running VMs..."
    $runningNames = @()

    try {
        $output = & $VBoxManagePath list runningvms 2>&1
        foreach ($line in $output) {
            if ($line -match '^"(.+?)"\s+\{') {
                $runningNames += $Matches[1]
            }
        }
        if ($runningNames.Count -eq 0) {
            Write-Log "INFO" "  No VMs currently running."
        } else {
            Write-Log "INFO" "  Running VMs: $($runningNames -join ', ')"
        }
    } catch {
        Write-Log "ERROR" "  VBoxManage query failed: $_"
        Set-OverallStatus "WARNING"
    }

    return $runningNames
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                      DOCUMENTS BACKUP
# =============================================================================

function Invoke-DocumentsBackup {
    Write-Log "INFO" "====================================================================="
    Write-Log "INFO" "  DOCUMENTS BACKUP  (version: $($script:RunTimestamp))"
    Write-Log "INFO" "====================================================================="

    Write-Log "INFO" "-- Documents to D: and E: (parallel) --"
    $versionedPathD = Get-VersionedPath -BaseDir $DestD -SourceName "Documents"
    $versionedPathE = Get-VersionedPath -BaseDir $DestE -SourceName "Documents"

    try {
        Invoke-ParallelLocalCopy `
            -SourceDir $SourceDocuments `
            -DestDirD  $versionedPathD  `
            -DestDirE  $versionedPathE  `
            -Label     "Documents"
    } catch {
        $msg = "Documents parallel copy FAILED -- $_"
        Write-Log "ERROR" $msg
        $script:ErrorsList.Add($msg)
        Set-OverallStatus "FAILURE"
    }

    # Prune old versions on each destination (sequential — lightweight metadata operations)
    try {
        Remove-OldVersions -BaseDir $DestD -SourceName "Documents" -KeepCount $DocumentVersionsToRetain
    } catch {
        Write-Log "WARNING" "Version pruning on D: failed -- $_"
    }
    try {
        Remove-OldVersions -BaseDir $DestE -SourceName "Documents" -KeepCount $DocumentVersionsToRetain
    } catch {
        Write-Log "WARNING" "Version pruning on E: failed -- $_"
    }

    Write-Log "INFO" "Documents to D: and E: complete."
    Write-Log "INFO" "-- Documents to Backblaze B2 --"
    Invoke-DocumentsB2Backup
}

function Invoke-DocumentsB2Backup {
    if (-not (Test-Path $TempCompressDir)) {
        New-Item -ItemType Directory -Path $TempCompressDir -Force | Out-Null
    }

    $archiveName = "Documents_$($script:RunTimestamp).7z"
    $archivePath = Join-Path $TempCompressDir $archiveName

    try {
        Write-Log "INFO" "  Compressing Documents to $archiveName"
        $szArgs   = @("a", "-t7z", "-mx=5", "-mmt=on", $archivePath, "$SourceDocuments\*")
        $szOutput = & $SevenZipPath @szArgs 2>&1
        Write-Log "INFO" "  7-Zip: $($szOutput | Select-Object -Last 3 | Out-String)".Trim()

        # 7-Zip exit codes:  0 = success, 1 = warning (e.g. access denied on a junction
        # point or locked file — archive is still usable), 2+ = fatal error.
        # We treat code 1 as a WARNING and continue; codes 2+ abort as FAILURE.
        if ($LASTEXITCODE -eq 1) {
            $warnMsg = "7-Zip completed with warnings (exit code 1) compressing Documents -- some files may have been skipped. Check the log above for details."
            Write-Log "WARNING" "  $warnMsg"
            $script:ErrorsList.Add($warnMsg)
            Set-OverallStatus "WARNING"
        } elseif ($LASTEXITCODE -gt 1) {
            throw "7-Zip exited with fatal code $LASTEXITCODE"
        }
        Write-Log "INFO" "  Compression complete: $archivePath"

        $b2Dest = "${RcloneRemoteName}:${B2BucketName}/Documents/"
        Write-Log "INFO" "  Uploading $archiveName to $b2Dest"
        $rcArgs = @("copy", "--checksum", "--stats-one-line", $archivePath, $b2Dest)
        $rcOut  = & $RclonePath @rcArgs 2>&1
        Write-Log "INFO" "  rclone: $($rcOut | Out-String)".Trim()

        if ($LASTEXITCODE -ne 0) { throw "rclone exited with code $LASTEXITCODE. Output: $rcOut" }
        Write-Log "INFO" "  Upload complete."

        Remove-OldB2Versions -RemotePath $b2Dest -FilePrefix "Documents_" -KeepCount $DocumentVersionsToRetain

    } catch {
        $msg = "Documents to B2 FAILED -- $_"
        Write-Log "ERROR" $msg
        $script:ErrorsList.Add($msg)
        Set-OverallStatus "FAILURE"
    } finally {
        if (Test-Path $archivePath) {
            Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
            Write-Log "INFO" "  Temp archive removed: $archiveName"
        }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                      VIRTUALBOX VMs BACKUP
# =============================================================================

function Invoke-VMBackup {
    Write-Log "INFO" "====================================================================="
    Write-Log "INFO" "  VIRTUALBOX VM BACKUP"
    Write-Log "INFO" "====================================================================="

    $runningVMs = Get-RunningVMNames

    $vmDirs = Get-ChildItem -Path $SourceVMs -Directory -ErrorAction SilentlyContinue
    if ($null -eq $vmDirs -or @($vmDirs).Count -eq 0) {
        Write-Log "WARNING" "No VM directories found in $SourceVMs"
        return
    }

    foreach ($vmDir in $vmDirs) {
        $vmName = $vmDir.Name

        if ($runningVMs -contains $vmName) {
            $msg = "SKIPPING VM '$vmName' -- currently running (unsafe to back up)."
            Write-Log "WARNING" $msg
            $script:SkippedVMs.Add($vmName)
            Set-OverallStatus "WARNING"
            continue
        }

        Write-Log "INFO" "Processing VM: $vmName"
        $script:BackedUpVMs.Add($vmName)

        Write-Log "INFO" "  VM '$vmName' to D: and E: (parallel)"
        try {
            Invoke-ParallelLocalCopy `
                -SourceDir $vmDir.FullName                          `
                -DestDirD  (Join-Path $DestD "VirtualBox VMs\$vmName") `
                -DestDirE  (Join-Path $DestE "VirtualBox VMs\$vmName") `
                -Label     "VM '$vmName'"
        } catch {
            $msg = "VM '$vmName' parallel local copy FAILED -- $_"
            Write-Log "ERROR" $msg; $script:ErrorsList.Add($msg); Set-OverallStatus "FAILURE"
        }

        Write-Log "INFO" "  VM '$vmName' to B2:"
        try {
            Invoke-VMB2Backup -VMDir $vmDir.FullName -VMName $vmName
        } catch {
            $msg = "VM '$vmName' to B2: FAILED -- $_"
            Write-Log "ERROR" $msg; $script:ErrorsList.Add($msg); Set-OverallStatus "FAILURE"
        }
    }
}

function Invoke-VMB2Backup {
    param([string]$VMDir, [string]$VMName)
    if (-not (Test-Path $TempCompressDir)) {
        New-Item -ItemType Directory -Path $TempCompressDir -Force | Out-Null
    }

    $safeName    = $VMName -replace '[\\/:*?"<>|]', '_'
    $archiveName = "VM_${safeName}_$($script:RunTimestamp).7z"
    $archivePath = Join-Path $TempCompressDir $archiveName

    try {
        Write-Log "INFO" "    Compressing '$VMName' to $archiveName"
        $szArgs   = @("a", "-t7z", "-mx=1", "-mmt=on", $archivePath, "$VMDir\*")
        $szOutput = & $SevenZipPath @szArgs 2>&1
        Write-Log "INFO" "    7-Zip: $($szOutput | Select-Object -Last 3 | Out-String)".Trim()

        # Exit code 1 = warning (skipped/locked file) — treat as WARNING, not FAILURE.
        # Exit codes 2+ = fatal error — abort.
        if ($LASTEXITCODE -eq 1) {
            $warnMsg = "7-Zip completed with warnings (exit code 1) compressing VM '$VMName' -- some files may have been skipped."
            Write-Log "WARNING" "    $warnMsg"
            $script:ErrorsList.Add($warnMsg)
            Set-OverallStatus "WARNING"
        } elseif ($LASTEXITCODE -gt 1) {
            throw "7-Zip exited with fatal code $LASTEXITCODE"
        }

        $b2Dest = "${RcloneRemoteName}:${B2BucketName}/VirtualBox VMs/"
        Write-Log "INFO" "    Uploading $archiveName to $b2Dest"
        $rcArgs = @("copy", "--checksum", "--stats-one-line", $archivePath, $b2Dest)
        $rcOut  = & $RclonePath @rcArgs 2>&1
        Write-Log "INFO" "    rclone: $($rcOut | Out-String)".Trim()

        if ($LASTEXITCODE -ne 0) { throw "rclone exited with code $LASTEXITCODE. Output: $rcOut" }
        Write-Log "INFO" "    VM '$VMName' uploaded to B2 successfully."

    } finally {
        if (Test-Path $archivePath) {
            Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
            Write-Log "INFO" "    Temp archive removed: $archiveName"
        }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                           FIRETEST
# =============================================================================

function Invoke-Firetest {
    Write-Log "INFO" "====================================================================="
    Write-Log "INFO" "  FIRETEST (Restore Verification)"
    Write-Log "INFO" "====================================================================="

    $vmExtensions = @('.vdi','.vmdk','.vhd','.vhdx','.vmx','.nvram','.vbox','.sav')

    foreach ($destBase in @($DestD, $DestE)) {
        Write-Log "INFO" "-- Firetest on $destBase --"

        $versDirs = Get-ChildItem -Path $destBase -Directory -Filter "Documents_????-??-??_??-??-??" -ErrorAction SilentlyContinue |
                    Sort-Object Name -Descending

        if ($null -eq $versDirs -or @($versDirs).Count -eq 0) {
            $note = "SKIP [$destBase] -- No Documents versioned backup found."
            Write-Log "WARNING" "  $note"
            $script:FiretestResults.Add($note)
            continue
        }

        $latestDir = $versDirs[0].FullName
        Write-Log "INFO" "  Using backup directory: $($versDirs[0].Name)"

        $candidates = Get-ChildItem -Path $latestDir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Extension.ToLower() -notin $vmExtensions -and
                $_.Length -ge $FiretestMinSizeBytes -and
                $_.Length -le $FiretestMaxSizeBytes
            }

        if ($null -eq $candidates -or @($candidates).Count -eq 0) {
            $note = "SKIP [$destBase] -- No eligible files found (size: $($FiretestMinSizeBytes/1KB)KB to $($FiretestMaxSizeBytes/1MB)MB, non-VM)."
            Write-Log "WARNING" "  $note"
            $script:FiretestResults.Add($note)
            continue
        }

        $chosen = @($candidates) | Get-Random
        Write-Log "INFO" "  Selected file : $($chosen.FullName)"
        Write-Log "INFO" "  File size     : $([math]::Round($chosen.Length / 1KB, 1)) KB"

        $relPath    = $chosen.FullName.Substring($latestDir.TrimEnd('\').Length).TrimStart('\')
        $sourceFile = Join-Path $SourceDocuments $relPath

        if (-not (Test-Path -LiteralPath $sourceFile)) {
            $note = "FAIL [$destBase] -- Source file not found for firetest: $sourceFile"
            Write-Log "ERROR" "  $note"
            $script:FiretestResults.Add($note)
            $script:ErrorsList.Add($note)
            Set-OverallStatus "FAILURE"
            continue
        }

        try {
            $destHash   = Get-FileHash256 $chosen.FullName
            $sourceHash = Get-FileHash256 $sourceFile

            if ($destHash -eq $sourceHash) {
                $note = "PASS [$destBase] | File: $relPath | Hash: $destHash"
                Write-Log "INFO" "  Firetest PASSED."
                Write-Log "INFO" "  $note"
                $script:FiretestResults.Add($note)
            } else {
                $note = "FAIL [$destBase] | File: $relPath | SrcHash: $sourceHash | DstHash: $destHash"
                Write-Log "ERROR" "  Firetest FAILED -- hash mismatch!"
                Write-Log "ERROR" "  $note"
                $script:FiretestResults.Add($note)
                $script:ErrorsList.Add($note)
                Set-OverallStatus "FAILURE"
            }
        } catch {
            $note = "ERROR [$destBase] | File: $relPath | Error: $_"
            Write-Log "ERROR" "  Firetest error: $_"
            $script:FiretestResults.Add($note)
            $script:ErrorsList.Add($note)
            Set-OverallStatus "FAILURE"
        }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#
#  DKIM SIGNING  (RFC 6376  --  relaxed/relaxed canonicalisation, rsa-sha256)
#
#  Key storage: Windows Certificate Store
#  ──────────────────────────────────────
#  The RSA-2048 key pair lives in Cert:\<$DkimCertStore>\My, identified by
#  $DkimCertSubject ("CN=BackupDKIM").  The certificate is self-signed
#  and created with KeyExportPolicy=NonExportable — the private key material
#  is held in the CNG Key Storage Provider and never written to disk as a
#  readable file.
#
#  PowerShell version compatibility
#  ─────────────────────────────────
#  New-SelfSignedCertificate on modern Windows always creates CNG-backed keys.
#  The RSA object obtained from the cert store via GetRSAPrivateKey() is
#  therefore always RSACng on any currently supported Windows version.
#  However, to be defensive, the signing helper detects the concrete type
#  at runtime and calls the appropriate SignData overload, so the script
#  also works on older machines where the legacy CSP might be in use.
#
#  DER helpers
#  ────────────
#  Used only to build the SubjectPublicKeyInfo blob for the DNS TXT p= tag.
#  When exporting the public key we call rsa.ExportParameters($false),
#  which works on both RSACng and RSACryptoServiceProvider and returns only
#  the public components regardless of the NonExportable policy on the
#  private key.
#
# =============================================================================

# ── DER encoding helpers ──────────────────────────────────────────────────────

function ConvertTo-DerLength {
    <#
    .SYNOPSIS  Encodes an integer as a BER/DER length field (short or long form).
    #>
    param([int]$Len)
    if ($Len -lt 128) {
        return [byte[]]@([byte]$Len)
    } elseif ($Len -lt 256) {
        return [byte[]]@([byte]0x81, [byte]$Len)
    } else {
        return [byte[]]@(
            [byte]0x82,
            [byte](($Len -shr 8) -band 0xFF),
            [byte]($Len -band 0xFF)
        )
    }
}

function ConvertTo-DerInteger {
    <#
    .SYNOPSIS  Encodes a byte array as a DER INTEGER, prepending 0x00 when the
               most-significant bit is set so the value is treated as positive.
    #>
    param([byte[]]$Bytes)
    if ($Bytes[0] -band 0x80) {
        $Bytes = @([byte]0x00) + $Bytes
    }
    return [byte[]](@([byte]0x02) + (ConvertTo-DerLength $Bytes.Length) + $Bytes)
}

function ConvertTo-DerSequence {
    <#
    .SYNOPSIS  Wraps a byte array in a DER SEQUENCE tag (0x30).
    #>
    param([byte[]]$Content)
    return [byte[]](@([byte]0x30) + (ConvertTo-DerLength $Content.Length) + $Content)
}

# ── SubjectPublicKeyInfo export ───────────────────────────────────────────────

function Export-DkimPublicKeyBase64 {
    <#
    .SYNOPSIS  Derives the base64-encoded SubjectPublicKeyInfo (SPKI) DER blob
               from an RSA object and returns it for use in the DKIM DNS p= tag.

    .PARAMETER RsaKey
        Any RSA object that supports ExportParameters($false).
        Compatible with RSACng (PS7 / CNG store) and RSACryptoServiceProvider
        (PS5.1 / legacy CSP).

    .NOTES
        Structure (RFC 5480 / X.509):
          SubjectPublicKeyInfo ::= SEQUENCE {
            algorithm  AlgorithmIdentifier  -- OID rsaEncryption + NULL
            subjectPublicKey  BIT STRING {
              RSAPublicKey ::= SEQUENCE { INTEGER(n), INTEGER(e) }
            }
          }
    #>
    param([System.Security.Cryptography.RSA]$RsaKey)

    $params = $RsaKey.ExportParameters($false)  # public components only

    # Inner RSAPublicKey (PKCS#1): SEQUENCE { INTEGER(n), INTEGER(e) }
    $rsaPublicKey = ConvertTo-DerSequence (
        (ConvertTo-DerInteger $params.Modulus) +
        (ConvertTo-DerInteger $params.Exponent)
    )

    # BIT STRING: tag(0x03) + length + unusedBits(0x00) + RSAPublicKey
    $bitStringPayload = @([byte]0x00) + $rsaPublicKey
    $bitString = [byte[]](@([byte]0x03) +
        (ConvertTo-DerLength $bitStringPayload.Length) +
        $bitStringPayload)

    # AlgorithmIdentifier for rsaEncryption (OID 1.2.840.113549.1.1.1) + NULL
    $algoIdentifierBytes = [byte[]]@(
        0x06, 0x09,
        0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x01, 0x01,
        0x05, 0x00
    )
    $algoIdentifier = ConvertTo-DerSequence $algoIdentifierBytes

    # SubjectPublicKeyInfo: SEQUENCE { AlgorithmIdentifier, BIT STRING }
    $spki = ConvertTo-DerSequence ($algoIdentifier + $bitString)

    return [Convert]::ToBase64String($spki)
}

# ── Certificate store helpers ─────────────────────────────────────────────────

function Get-DkimCertificate {
    <#
    .SYNOPSIS  Retrieves the DKIM certificate from the configured cert store.
               Throws a descriptive error if the certificate is not found.
    #>
    $storePath = "Cert:\${DkimCertStore}\My"
    $cert = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
            Where-Object { $_.Subject -eq $DkimCertSubject } |
            Select-Object -First 1

    if (-not $cert) {
        throw "DKIM certificate '$DkimCertSubject' not found in $storePath. " +
              "Run Initialize-DkimKeys to generate it."
    }
    return $cert
}

function Get-RsaPrivateKeyFromCert {
    <#
    .SYNOPSIS  Returns an RSA object backed by the certificate's private key.

    .NOTES
        GetRSAPrivateKey() is available in .NET Framework 4.6+ (PS 5.1 on
        Windows 10/11) and all .NET 5+ versions.  It returns RSACng for CNG-
        stored keys and RSACryptoServiceProvider for legacy CSP keys.

        The caller must NOT dispose the returned object — for RSACng the
        lifetime is tied to the certificate object.
    #>
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)

    $rsaKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)

    if (-not $rsaKey) {
        throw "Could not obtain RSA private key from certificate '$($Cert.Subject)'. " +
              "Verify the Task Scheduler account has permission to access this certificate's " +
              "private key (certmgr.msc / certlm.msc -> All Tasks -> Manage Private Keys)."
    }
    return $rsaKey
}

function Invoke-RsaSha256Sign {
    <#
    .SYNOPSIS  Signs $DataBytes using RSA-SHA256 PKCS#1 v1.5.
               Dispatches to the correct SignData overload based on the
               concrete RSA type to maintain PS 5.1 / PS 7+ compatibility.

    .OUTPUTS   [byte[]] Raw RSA signature bytes.
    #>
    param(
        [System.Security.Cryptography.RSA]$RsaKey,
        [byte[]]$DataBytes
    )

    # RSACng path — used on modern Windows (CNG-backed keys, PS 5.1 + .NET 4.6+ and PS 7+)
    if ($RsaKey -is [System.Security.Cryptography.RSACng]) {
        return $RsaKey.SignData(
            $DataBytes,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    }

    # RSACryptoServiceProvider path — legacy CSP, PS 5.1 on older hardware
    if ($RsaKey -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
        $sha256 = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
        try {
            return $RsaKey.SignData($DataBytes, $sha256)
        } finally {
            $sha256.Dispose()
        }
    }

    # Fallback for any other RSA implementation (e.g. custom provider)
    # Uses the abstract RSA.SignData overload available in .NET 4.6+
    try {
        return $RsaKey.SignData(
            $DataBytes,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    } catch {
        throw "Unsupported RSA key type '$($RsaKey.GetType().FullName)': $_"
    }
}

# ── Key generation ────────────────────────────────────────────────────────────

function Initialize-DkimKeys {
    <#
    .SYNOPSIS  Creates a NonExportable RSA-2048 DKIM certificate in the Windows
               Certificate Store if one does not already exist.

               On first run:
               - Creates a self-signed certificate ($DkimCertSubject) in
                 Cert:\<$DkimCertStore>\My with a 10-year validity.
               - Prints the Cloudflare DNS TXT record you must add.
               - Logs the record and the certificate thumbprint.

               On subsequent runs:
               - Detects the existing certificate and returns immediately.

    .NOTES
        NonExportable means the private key cannot be extracted via PFX export.
        If the machine is decommissioned you will need to generate a new key
        pair and update the DNS record.  The certificate thumbprint is logged
        so you can identify it in certmgr.msc or certlm.msc.

        To verify DNS propagation after adding the record:
          nslookup -type=TXT <selector>._domainkey.<yourdomain.com> 1.1.1.1
    #>

    $storePath = "Cert:\${DkimCertStore}\My"

    $existing = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
                Where-Object { $_.Subject -eq $DkimCertSubject } |
                Select-Object -First 1

    if ($existing) {
        Write-Log "INFO" "DKIM certificate already present in ${storePath}."
        Write-Log "INFO" "  Subject    : $($existing.Subject)"
        Write-Log "INFO" "  Thumbprint : $($existing.Thumbprint)"
        Write-Log "INFO" "  Expires    : $($existing.NotAfter.ToString('yyyy-MM-dd'))"
        return
    }

    Write-Log "INFO" "DKIM certificate not found -- generating RSA-2048 key pair in ${storePath}..."

    try {
        # New-SelfSignedCertificate is part of the PKI module (available on
        # Windows 8.1+ / Server 2012 R2+, always present on Windows 10/11).
        # KeyExportPolicy = NonExportable ensures the private key material
        # cannot be exported as a PFX or PEM file.
        $cert = New-SelfSignedCertificate `
            -Subject          $DkimCertSubject       `
            -KeyAlgorithm     RSA                    `
            -KeyLength        2048                   `
            -KeyExportPolicy  NonExportable          `
            -KeyUsage         None                   `
            -CertStoreLocation $storePath            `
            -NotAfter         (Get-Date).AddYears(10) `
            -ErrorAction      Stop

        Write-Log "INFO" "Certificate created.  Thumbprint: $($cert.Thumbprint)"

        # Derive the public key from the cert for the DNS record.
        # ExportParameters($false) exports only public components and is
        # permitted even when the private key is NonExportable.
        $rsaPub  = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)
        $pubB64  = Export-DkimPublicKeyBase64 -RsaKey $rsaPub

        $dnsName  = "${DkimSelector}._domainkey.${DkimDomain}"
        $dnsValue = "v=DKIM1; k=rsa; p=${pubB64}"

        Write-Log "INFO" "=========================================================="
        Write-Log "INFO" "  ADD THIS DNS TXT RECORD IN CLOUDFLARE"
        Write-Log "INFO" "  Name  : $dnsName"
        Write-Log "INFO" "  Type  : TXT"
        Write-Log "INFO" "  TTL   : Auto (300)"
        Write-Log "INFO" "  Value : $dnsValue"
        Write-Log "INFO" "=========================================================="

        Write-Host ""
        Write-Host "==========================================================" -ForegroundColor Green
        Write-Host "  DKIM KEY GENERATED -- ADD THIS DNS TXT RECORD NOW"       -ForegroundColor Green
        Write-Host "==========================================================" -ForegroundColor Green
        Write-Host "  Name  : $dnsName"  -ForegroundColor Yellow
        Write-Host "  Type  : TXT"        -ForegroundColor Yellow
        Write-Host "  TTL   : Auto"       -ForegroundColor Yellow
        Write-Host "  Value : $dnsValue"  -ForegroundColor Yellow
        Write-Host "==========================================================" -ForegroundColor Green
        Write-Host "  Certificate thumbprint : $($cert.Thumbprint)"            -ForegroundColor Cyan
        Write-Host "  Store                  : $storePath"                     -ForegroundColor Cyan
        Write-Host "  Verify DNS propagation :"                                -ForegroundColor Cyan
        Write-Host "    nslookup -type=TXT $dnsName 1.1.1.1"                  -ForegroundColor Cyan
        Write-Host "==========================================================" -ForegroundColor Green
        Write-Host ""

    } catch {
        throw "Failed to create DKIM certificate in ${storePath}: $_"
    }
}

# ── Canonicalisation ──────────────────────────────────────────────────────────

function Get-DkimBodyHash {
    <#
    .SYNOPSIS  RFC 6376 section 3.4.4 relaxed body canonicalisation.
               Returns base64-encoded SHA-256 of the canonical body.
    #>
    param([string]$BodyText)

    $text  = $BodyText -replace '\r\n', "`n" -replace '\r', "`n"
    $lines = ($text -split "`n") | ForEach-Object {
        ($_ -replace '[ \t]+', ' ') -replace '[ \t]+$', ''
    }
    $canonical = ($lines -join "`r`n").TrimEnd("`r`n") + "`r`n"

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToBase64String(
            $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($canonical))
        )
    } finally {
        $sha256.Dispose()
    }
}

function Get-RelaxedHeaderLine {
    <#
    .SYNOPSIS  RFC 6376 section 3.4.2 relaxed header canonicalisation.
               Returns "lowercase-name:normalised-value" without trailing CRLF.
    #>
    param([string]$Name, [string]$Value)
    $canonName  = $Name.ToLower().Trim()
    $canonValue = ($Value -replace '\r\n[ \t]+', ' ') -replace '[ \t]+', ' '
    $canonValue = $canonValue.Trim()
    return "${canonName}:${canonValue}"
}

# ── Signature construction ────────────────────────────────────────────────────

function New-DkimSignatureHeader {
    <#
    .SYNOPSIS  Builds a complete, signed DKIM-Signature header value.

    .PARAMETER SignedHeaders
        Hashtable mapping lowercase header name to raw value for every header
        in the h= tag.

    .PARAMETER HeaderOrder
        Ordered array of lowercase header names (defines h= and signing sequence).

    .PARAMETER BodyHash
        Base64 SHA-256 body hash (bh= tag value).

    .OUTPUTS   [string]  Full DKIM-Signature header value (no "DKIM-Signature: " prefix).
    #>
    param(
        [hashtable] $SignedHeaders,
        [string[]]  $HeaderOrder,
        [string]    $BodyHash
    )

    $cert   = Get-DkimCertificate
    $rsaKey = Get-RsaPrivateKeyFromCert -Cert $cert

    $hTag      = ($HeaderOrder | ForEach-Object { $_.ToLower() }) -join ':'
    $timestamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    # Partial DKIM-Signature value with b= empty — filled in after signing
    $partialVal = "v=1; a=rsa-sha256; c=relaxed/relaxed; " +
                  "d=${DkimDomain}; s=${DkimSelector}; " +
                  "t=${timestamp}; h=${hTag}; " +
                  "bh=${BodyHash}; b="

    # Build signing data (RFC 6376 section 3.7):
    #   canonical(header_1) CRLF ... canonical(header_N) CRLF
    #   canonical("dkim-signature:" + partialVal)   -- NO trailing CRLF
    $sb = [System.Text.StringBuilder]::new()
    foreach ($hName in $HeaderOrder) {
        $sb.Append((Get-RelaxedHeaderLine $hName $SignedHeaders[$hName]) + "`r`n") | Out-Null
    }
    $sb.Append("dkim-signature:${partialVal}") | Out-Null

    $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $sigBytes  = Invoke-RsaSha256Sign -RsaKey $rsaKey -DataBytes $dataBytes

    return $partialVal + [Convert]::ToBase64String($sigBytes)
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#
#  RAW SMTP TRANSPORT  (TCP + STARTTLS, AUTH LOGIN)
#
#  The complete RFC 5322 message is constructed first, signed with DKIM,
#  and then delivered verbatim so that the receiving MTA verifies exactly
#  the bytes that were signed.
#
# =============================================================================

function Read-SmtpResponse {
    param([System.IO.StreamReader]$Reader)
    $lines = [System.Collections.Generic.List[string]]::new()
    do {
        $line = $Reader.ReadLine()
        if ($null -eq $line) { break }
        $lines.Add($line)
    } while ($line.Length -ge 4 -and $line[3] -eq '-')
    return $lines.ToArray()
}

function Invoke-SmtpCommand {
    param(
        [System.IO.StreamWriter] $Writer,
        [System.IO.StreamReader] $Reader,
        [AllowNull()][string]    $Command,
        [int]                    $ExpectedCode
    )
    if (-not [string]::IsNullOrEmpty($Command)) { $Writer.WriteLine($Command) }
    # @() forces the return value to remain a [string[]] array even when the server
    # sends a single-line response.  Without this, PowerShell unwraps a one-element
    # array to a bare string, making $resp[-1] return a [char] instead of a [string],
    # and the subsequent .Substring(0,3) call throws a method-not-found error.
    $resp     = @(Read-SmtpResponse -Reader $Reader)
    $lastLine = $resp[-1]
    $code     = [int]($lastLine.Substring(0, 3))
    if ($code -ne $ExpectedCode) {
        throw "SMTP: expected $ExpectedCode, received $code -- $lastLine"
    }
    return $resp
}

function Send-DkimSignedEmail {
    <#
    .SYNOPSIS  Builds a DKIM-signed RFC 5322 message and delivers it via raw
               SMTP with optional STARTTLS.  On failure the error is logged
               without escalating $script:OverallStatus.
    #>
    param([string]$Subject, [string]$Body)

    $tcp       = $null
    $netStream = $null
    $sslStream = $null

    try {
        # ── 1. RFC 5322 header values ─────────────────────────────────────

        $now       = Get-Date
        $utcOffset = [System.TimeZoneInfo]::Local.GetUtcOffset($now)
        $tzSign    = if ($utcOffset.TotalMinutes -ge 0) { '+' } else { '-' }
        $tzStr     = '{0}{1:D2}{2:D2}' -f $tzSign, [math]::Abs($utcOffset.Hours), [math]::Abs($utcOffset.Minutes)
        $dateValue = $now.ToString("ddd, dd MMM yyyy HH:mm:ss") + " $tzStr"

        $msgIdValue   = "<backup-$(Get-Date -Format 'yyyyMMddHHmmss')-$(Get-Random -Maximum 99999)@${DkimDomain}>"
        $fromValue    = "${EmailFromName} <${EmailFrom}>"
        $toValue      = $EmailTo
        $subjectValue = $Subject

        $headerOrder   = @('from', 'to', 'subject', 'date', 'message-id')
        $signedHeaders = @{
            'from'       = $fromValue
            'to'         = $toValue
            'subject'    = $subjectValue
            'date'       = $dateValue
            'message-id' = $msgIdValue
        }

        # ── 2. Body hash ──────────────────────────────────────────────────
        $bodyHash = Get-DkimBodyHash -BodyText $Body

        # ── 3. DKIM signature ─────────────────────────────────────────────
        $dkimSigValue = New-DkimSignatureHeader `
            -SignedHeaders $signedHeaders `
            -HeaderOrder   $headerOrder   `
            -BodyHash      $bodyHash

        # ── 4. Assemble raw RFC 5322 message ──────────────────────────────
        $txBody  = ($Body -replace '\r\n', "`n" -replace '\r', "`n") -replace "`n", "`r`n"
        $txBody  = $txBody.TrimEnd("`r`n") + "`r`n"

        $rawMsg  = "DKIM-Signature: ${dkimSigValue}`r`n"
        $rawMsg += "From: ${fromValue}`r`n"
        $rawMsg += "To: ${toValue}`r`n"
        $rawMsg += "Subject: ${subjectValue}`r`n"
        $rawMsg += "Date: ${dateValue}`r`n"
        $rawMsg += "Message-ID: ${msgIdValue}`r`n"
        $rawMsg += "MIME-Version: 1.0`r`n"
        $rawMsg += "Content-Type: text/plain; charset=UTF-8`r`n"
        $rawMsg += "Content-Transfer-Encoding: 8bit`r`n"
        $rawMsg += "`r`n"
        $rawMsg += $txBody

        # ── 5. Dot-stuffing (RFC 5321 section 4.5.2) ─────────────────────
        $msgLines = $rawMsg -split "`r`n"
        while ($msgLines.Count -gt 0 -and $msgLines[-1] -eq '') {
            $msgLines = $msgLines[0..($msgLines.Length - 2)]
        }
        $stuffedBlock = (
            ($msgLines | ForEach-Object { if ($_.StartsWith('.')) { ".$_" } else { $_ } }) -join "`r`n"
        ) + "`r`n"

        # ── 6. SMTP conversation ──────────────────────────────────────────
        # UTF8Encoding($false) = no BOM.  [System.Text.Encoding]::UTF8 emits a
        # UTF-8 BOM (0xEF 0xBB 0xBF) before the first StreamWriter.Write() call
        # on .NET Framework (PowerShell 5.1).  The BOM prepended to the first
        # EHLO command makes it unrecognisable to the server, returning 500.
        $enc = New-Object System.Text.UTF8Encoding($false)

        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect($SmtpServer, $SmtpPort)
        $netStream = $tcp.GetStream()

        if ($SmtpImplicitTls) {
            # ── Implicit TLS (port 465) ───────────────────────────────────
            # TLS handshake occurs immediately on TCP connect — before the
            # server sends its greeting.  Wrap the raw stream in SslStream
            # first, then bind reader/writer and start the SMTP conversation.
            $tlsHost   = if ($SmtpTlsHostname -ne '') { $SmtpTlsHostname } else { $SmtpServer }
            $sslStream = New-Object System.Net.Security.SslStream($netStream, $false)
            $sslStream.AuthenticateAsClient($tlsHost)

            $reader = New-Object System.IO.StreamReader($sslStream, $enc)
            $writer = New-Object System.IO.StreamWriter($sslStream, $enc)
            $writer.NewLine   = "`r`n"
            $writer.AutoFlush = $true

            Invoke-SmtpCommand $writer $reader $null                  220 | Out-Null  # greeting (over TLS)
            Invoke-SmtpCommand $writer $reader "EHLO ${DkimDomain}"   250 | Out-Null

        } else {
            # ── STARTTLS (port 587) ───────────────────────────────────────
            # Connect plain, read greeting, send EHLO, issue STARTTLS, then
            # upgrade the stream and re-send EHLO over TLS (RFC 3207 §4).
            $reader = New-Object System.IO.StreamReader($netStream, $enc)
            $writer = New-Object System.IO.StreamWriter($netStream, $enc)
            $writer.NewLine   = "`r`n"
            $writer.AutoFlush = $true

            Invoke-SmtpCommand $writer $reader $null                  220 | Out-Null  # plain greeting
            Invoke-SmtpCommand $writer $reader "EHLO ${DkimDomain}"   250 | Out-Null
            Invoke-SmtpCommand $writer $reader "STARTTLS"             220 | Out-Null

            $tlsHost   = if ($SmtpTlsHostname -ne '') { $SmtpTlsHostname } else { $SmtpServer }
            $sslStream = New-Object System.Net.Security.SslStream($netStream, $false)
            $sslStream.AuthenticateAsClient($tlsHost)

            $reader = New-Object System.IO.StreamReader($sslStream, $enc)
            $writer = New-Object System.IO.StreamWriter($sslStream, $enc)
            $writer.NewLine   = "`r`n"
            $writer.AutoFlush = $true

            Invoke-SmtpCommand $writer $reader "EHLO ${DkimDomain}"   250 | Out-Null  # re-EHLO after TLS
        }

        Invoke-SmtpCommand $writer $reader "AUTH LOGIN" 334 | Out-Null
        Invoke-SmtpCommand $writer $reader ([Convert]::ToBase64String($enc.GetBytes($SmtpUsername)))         334 | Out-Null
        Invoke-SmtpCommand $writer $reader ([Convert]::ToBase64String($enc.GetBytes($script:SmtpPassword))) 235 | Out-Null

        Invoke-SmtpCommand $writer $reader "MAIL FROM:<${EmailFrom}>" 250 | Out-Null
        Invoke-SmtpCommand $writer $reader "RCPT TO:<${EmailTo}>"     250 | Out-Null

        Invoke-SmtpCommand $writer $reader "DATA" 354 | Out-Null
        $writer.Write($stuffedBlock)
        $writer.WriteLine(".")
        Invoke-SmtpCommand $writer $reader $null 250 | Out-Null

        $writer.WriteLine("QUIT")

        Write-Log "INFO" "DKIM-signed email sent successfully. Subject: $Subject"

    } catch {
        Write-Log "ERROR" "Email delivery failed: $_"
    } finally {
        if ($sslStream)  { try { $sslStream.Close()  } catch {} }
        if ($netStream)  { try { $netStream.Close()  } catch {} }
        if ($tcp)        { try { $tcp.Close()        } catch {} }
    }
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                      EMAIL REPORT BUILDER
# =============================================================================

function Build-EmailBody {
    $sb  = [System.Text.StringBuilder]::new()
    $div = "=" * 60

    $elapsed    = (Get-Date) - $script:RunStartTime
    $elapsedStr = Format-Elapsed -Elapsed $elapsed

    $sb.AppendLine($div)                                                         | Out-Null
    $sb.AppendLine("  $BackupName BACKUP REPORT")                                | Out-Null
    $sb.AppendLine($div)                                                         | Out-Null
    $sb.AppendLine("  Timestamp  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")  | Out-Null
    $sb.AppendLine("  Status     : $($script:OverallStatus)")                   | Out-Null
    $sb.AppendLine("  Duration   : $elapsedStr")                                | Out-Null
    $sb.AppendLine("  Log File   : $($script:LogFile)")                         | Out-Null
    $sb.AppendLine($div)                                                         | Out-Null
    $sb.AppendLine("")                                                           | Out-Null

    $sb.AppendLine("DOCUMENTS")                                                  | Out-Null
    $sb.AppendLine("-" * 40)                                                     | Out-Null
    $sb.AppendLine("  Version stamp : $($script:RunTimestamp)")                 | Out-Null
    $sb.AppendLine("")                                                           | Out-Null

    $sb.AppendLine("VIRTUALBOX VMs")                                             | Out-Null
    $sb.AppendLine("-" * 40)                                                     | Out-Null
    if ($script:BackedUpVMs.Count -gt 0) {
        $sb.AppendLine("  Backed up         : $($script:BackedUpVMs -join ', ')") | Out-Null
    } else {
        $sb.AppendLine("  Backed up         : (none)")                          | Out-Null
    }
    if ($script:SkippedVMs.Count -gt 0) {
        $sb.AppendLine("  Skipped (running) : $($script:SkippedVMs -join ', ')") | Out-Null
    }
    $sb.AppendLine("")                                                           | Out-Null

    $sb.AppendLine("INTEGRITY CHECKS (Local Destinations)")                     | Out-Null
    $sb.AppendLine("-" * 40)                                                     | Out-Null
    $sb.AppendLine("  Passed : $($script:IntegrityPassed)")                     | Out-Null
    $sb.AppendLine("  Failed : $($script:IntegrityFailed)")                     | Out-Null
    $sb.AppendLine("")                                                           | Out-Null

    $sb.AppendLine("FIRETEST (Restore Verification)")                            | Out-Null
    $sb.AppendLine("-" * 40)                                                     | Out-Null
    foreach ($r in $script:FiretestResults) {
        $sb.AppendLine("  $r")                                                   | Out-Null
    }
    $sb.AppendLine("")                                                           | Out-Null

    if ($script:ErrorsList.Count -gt 0) {
        $sb.AppendLine("ERRORS AND WARNINGS")                                    | Out-Null
        $sb.AppendLine("-" * 40)                                                 | Out-Null
        foreach ($e in $script:ErrorsList) {
            $sb.AppendLine("  [!] $e")                                           | Out-Null
        }
        $sb.AppendLine("")                                                       | Out-Null
    }

    $sb.AppendLine($div)                                                         | Out-Null
    $sb.AppendLine("  End of report.")                                           | Out-Null
    $sb.AppendLine($div)                                                         | Out-Null

    return $sb.ToString()
}

function Get-EmailSubject {
    $prefix = switch ($script:OverallStatus) {
        "SUCCESS" { "[BACKUP OK]"      }
        "WARNING" { "[BACKUP WARNING]" }
        "FAILURE" { "[BACKUP FAILED]"  }
        default   { "[BACKUP]"         }
    }
    return "$prefix $BackupName -- $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
}

#endregion ===================================================================
###############################################################################


###############################################################################
#region ======================================================================
#                           MAIN ENTRY POINT
# =============================================================================

try {
    # ── 1. Initialise logging ─────────────────────────────────────────────
    Initialize-Log

    # ── 2. DKIM certificate check / first-time key generation ─────────────
    try {
        Initialize-DkimKeys
    } catch {
        Write-Log "WARNING" "DKIM key initialisation failed: $_"
    }

    # ── 3. Load SMTP password from Windows Credential Manager ─────────────
    #    Must happen before any email send (including the start notification).
    Get-SmtpPasswordFromCredentialManager

    # ── 4. Send backup-started notification ──────────────────────────────
    try {
        Send-BackupStartedEmail
    } catch {
        Write-Log "WARNING" "Backup-started email failed: $_"
    }

    # ── 5. Missed-schedule check ──────────────────────────────────────────
    try {
        Test-MissedSchedule
    } catch {
        Write-Log "WARNING" "Missed-schedule check failed: $_"
    }

    # ── 6. Validate source paths ──────────────────────────────────────────
    Write-Log "INFO" "--- Validating source paths ---"
    foreach ($src in @($SourceDocuments, $SourceVMs)) {
        if (Test-Path $src) {
            Write-Log "INFO" "  Source OK : $src"
        } else {
            $msg = "Source path not accessible: $src"
            Write-Log "ERROR" $msg
            $script:ErrorsList.Add($msg)
            Set-OverallStatus "FAILURE"
        }
    }

    # ── 7. Validate / create destination paths ────────────────────────────
    Write-Log "INFO" "--- Validating destination paths ---"
    foreach ($dest in @($DestD, $DestE)) {
        $drive = Split-Path $dest -Qualifier
        if (-not (Test-Path $drive)) {
            $msg = "Destination drive not accessible: $drive (needed for $dest)"
            Write-Log "ERROR" $msg
            $script:ErrorsList.Add($msg)
            Set-OverallStatus "FAILURE"
        } else {
            if (-not (Test-Path $dest)) {
                New-Item -ItemType Directory -Path $dest -Force | Out-Null
                Write-Log "INFO" "  Created destination: $dest"
            } else {
                Write-Log "INFO" "  Destination OK : $dest"
            }
        }
    }

    # ── 8. Dependency checks ──────────────────────────────────────────────
    Assert-Dependencies

    # ── 9. Documents backup (D, E, B2) ───────────────────────────────────
    Invoke-DocumentsBackup

    # ── 10. VirtualBox VM backup (D, E, B2) ──────────────────────────────
    Invoke-VMBackup

    # ── 11. Restore firetest ──────────────────────────────────────────────
    Invoke-Firetest

    # ── 12. Mark completion ───────────────────────────────────────────────
    $elapsed = Format-Elapsed -Elapsed ((Get-Date) - $script:RunStartTime)
    Write-Log "INFO" "Backup run completed with overall status: $($script:OverallStatus) (duration: $elapsed)"

} catch {
    $fatalMsg = "FATAL ERROR -- unhandled exception: $_"
    Write-Log "ERROR" $fatalMsg
    $script:ErrorsList.Add($fatalMsg)
    Set-OverallStatus "FAILURE"
} finally {
    # ── 13. Purge old logs ────────────────────────────────────────────────
    try { Remove-OldLogs } catch { Write-Log "WARNING" "Log purge error: $_" }

    # ── 14. Send DKIM-signed completion report ────────────────────────────
    $emailBody    = Build-EmailBody
    $emailSubject = Get-EmailSubject
    Send-DkimSignedEmail -Subject $emailSubject -Body $emailBody

    $finalElapsed = Format-Elapsed -Elapsed ((Get-Date) - $script:RunStartTime)
    Write-Log "INFO" "====================================================================="
    Write-Log "INFO" "  Windows Backup -- DONE  |  Status: $($script:OverallStatus)  |  Duration: $finalElapsed"
    Write-Log "INFO" "====================================================================="
}

#endregion ===================================================================
###############################################################################
