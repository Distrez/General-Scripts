Import-Module WebAdministration

# -----------------------------
# FUNCTION: Validate Credentials
# -----------------------------
function Test-UserCredentials {
    param(
        [string]$Username,
        [string]$Password
    )

    $signature = @"
using System;
using System.Runtime.InteropServices;

public class LogonChecker {
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool LogonUser(
        string lpszUsername,
        string lpszDomain,
        string lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        out IntPtr phToken
    );
}
"@

    Add-Type $signature -ErrorAction SilentlyContinue

    $token = [IntPtr]::Zero

    # Detect domain\user or .\user or user
    if ($Username -match "^(.*)\\(.*)$") {
        $domain = $Matches[1]
        $user   = $Matches[2]
    } else {
        $domain = $env:COMPUTERNAME
        $user   = $Username
    }

    # LOGON32_LOGON_NETWORK (3)
    # LOGON32_PROVIDER_DEFAULT (0)
    $result = [LogonChecker]::LogonUser(
        $user, $domain, $Password, 3, 0, [ref]$token
    )

    return $result
}


# -----------------------------
# GET USER INPUT
# -----------------------------
Write-Host "`n--- DocuWare App Pool Identity Update ---" -ForegroundColor Cyan

$username = Read-Host "Enter the username (Example: .\admin or domain\user)"

# Secure password prompt
$SecurePass = Read-Host "Enter the password" -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringUni(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePass)
            )

Write-Host "`n--- Checking Credentials ---`n" -ForegroundColor Cyan


# -----------------------------
# VALIDATE CREDENTIALS
# -----------------------------
if (-not (Test-UserCredentials -Username $username -Password $password)) {
    Write-Host "✗ ERROR: Invalid username or password. Stopping script!" -ForegroundColor Red
    Enter-Host "Press anything to exit"
	Return
}

Write-Host "✔ Credentials are valid. Proceeding..." -ForegroundColor Green


# -----------------------------
# GET DocuWare APP POOLS
# -----------------------------
$appPools = Get-ChildItem "IIS:\AppPools" |
            Where-Object { $_.Name -like "DocuWare*" }

Write-Host "`nFound $($appPools.Count) DocuWare App Pools:`n" -ForegroundColor Cyan
$appPools | Select-Object Name


# -----------------------------
# UPDATE EACH APP POOL
# -----------------------------
foreach ($pool in $appPools) {

    Write-Host "`nUpdating $($pool.Name)..." -ForegroundColor Yellow
    
    $demoPool = Get-Item ("IIS:\AppPools\" + $pool.Name)

    $demoPool.processModel.userName = $username
    $demoPool.processModel.password = $password
    $demoPool.processModel.identityType = 3   # SpecificUser

    $demoPool | Set-Item

    Write-Host "✔ Updated $($pool.Name)" -ForegroundColor Green
}

Write-Host "`n--- DONE ---`n" -ForegroundColor Cyan
