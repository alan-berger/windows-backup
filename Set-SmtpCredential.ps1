#Requires -Version 5.1

<#
.SYNOPSIS
    Stores the Windows Backup SMTP password securely in Windows Credential Manager.

.DESCRIPTION
    Run this script ONCE (or whenever the SMTP password changes) to store the
    password in Windows Credential Manager under the target name configured in
    WindowsBackup.ps1 ($SmtpCredentialTarget, default: "WindowsBackupSmtp").

    The password is protected by the Windows Data Protection API (DPAPI) and is
    only accessible to the Windows account that runs this script.  It is never
    written to disk as plaintext.

    WindowsBackup.ps1 retrieves the password at runtime using the same target
    name, so both scripts must use the same value.

.NOTES
    - Run this script as the SAME user account that will run WindowsBackup.ps1
      (i.e. the account configured in Task Scheduler).
    - If the task runs as SYSTEM, you must run this script as SYSTEM too — the
      credential is user-scoped.  The simplest way is to use PsExec:
          psexec -i -s powershell.exe Set-SmtpCredential.ps1
    - To update the password later, simply run this script again with the new
      password.  The existing credential is overwritten.
    - To delete the credential:
          cmdkey /delete:WindowsBackupSmtp
      or via Control Panel → Credential Manager → Windows Credentials.

.EXAMPLE
    .\Set-SmtpCredential.ps1

.EXAMPLE
    # Use a non-default target name (must match $SmtpCredentialTarget in the backup script)
    .\Set-SmtpCredential.ps1 -TargetName "WindowsBackupSmtpProd"
#>

[CmdletBinding()]
param(
    # Target name in Windows Credential Manager.
    # Must match $SmtpCredentialTarget in WindowsBackup.ps1.
    [string]$TargetName = "WindowsBackupSmtp",

    # The username (SMTP login address).  Pre-filled from the default config value;
    # change this if your SMTP username differs.
    [string]$Username   = "backup@yourdomain.com"
)

Set-StrictMode -Version Latest

###############################################################################
#  Load Win32 CredWrite via P/Invoke
#  Compatible with PowerShell 5.1 (.NET Framework 4.x) and PowerShell 7+.
###############################################################################

if (-not ([System.Management.Automation.PSTypeName]'WindowsBackup.CredentialWriter').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsBackup {
    public class CredentialWriter {

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
        private static extern bool CredWrite([In] ref CREDENTIAL cred, uint flags);

        // Writes a Generic (type=1) credential to Windows Credential Manager.
        // CRED_PERSIST_LOCAL_MACHINE (2) = persists across logon sessions on this machine.
        // The password is passed as a Unicode byte array; the pinned GCHandle
        // ensures the GC does not move the bytes before CredWrite reads them.
        // The array is zeroed immediately after CredWrite returns.
        public static void WritePassword(string target, string username, string password) {
            var bytes = Encoding.Unicode.GetBytes(password);
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try {
                var cred = new CREDENTIAL {
                    Type               = 1,    // CRED_TYPE_GENERIC
                    TargetName         = target,
                    UserName           = username,
                    CredentialBlob     = handle.AddrOfPinnedObject(),
                    CredentialBlobSize = (uint)bytes.Length,
                    Persist            = 2,    // CRED_PERSIST_LOCAL_MACHINE
                };
                if (!CredWrite(ref cred, 0))
                    throw new InvalidOperationException(
                        "CredWrite failed. Win32 error: " + Marshal.GetLastWin32Error());
            } finally {
                handle.Free();
                Array.Clear(bytes, 0, bytes.Length);   // zero the plaintext bytes
            }
        }
    }
}
'@ -ErrorAction Stop
}

###############################################################################
#  Main
###############################################################################

Write-Host ""
Write-Host "Windows Backup — SMTP Credential Setup" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Target name : $TargetName"  -ForegroundColor Yellow
Write-Host "  Username    : $Username"     -ForegroundColor Yellow
Write-Host ""
Write-Host "Enter the SMTP password (App Password if using Gmail)."
Write-Host "The password will NOT be echoed to the screen."
Write-Host ""

$securePassword = Read-Host -Prompt "SMTP password" -AsSecureString

# Verify it's not empty
$testBSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
    $testPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($testBSTR)
    if ([string]::IsNullOrEmpty($testPlain)) {
        Write-Host ""
        Write-Host "ERROR: Password cannot be empty." -ForegroundColor Red
        exit 1
    }
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($testBSTR)
}

# Convert SecureString to plain string for CredWrite
# The BSTR is zeroed immediately after conversion via ZeroFreeBSTR.
$bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
$plain = $null
try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    [WindowsBackup.CredentialWriter]::WritePassword($TargetName, $Username, $plain)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    # $plain is now eligible for GC; we cannot zero a .NET string's internal
    # buffer from PowerShell, but it existed in memory only for the CredWrite call.
}

Write-Host ""
Write-Host "SUCCESS: Credential stored in Windows Credential Manager." -ForegroundColor Green
Write-Host ""
Write-Host "  Target  : $TargetName"    -ForegroundColor Cyan
Write-Host "  Username: $Username"       -ForegroundColor Cyan
Write-Host ""
Write-Host "You can verify it in:" -ForegroundColor Cyan
Write-Host "  Control Panel --> Credential Manager --> Windows Credentials" -ForegroundColor Cyan
Write-Host ""
Write-Host "To delete it later:" -ForegroundColor Cyan
Write-Host "  cmdkey /delete:$TargetName" -ForegroundColor Cyan
Write-Host ""
