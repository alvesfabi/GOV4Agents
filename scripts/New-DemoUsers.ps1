<#
.SYNOPSIS
  Creates the GOV4Agents demo personas (sponsor / approver / manager) in a
  target tenant using device code flow.

.DESCRIPTION
  This script:
  1. Authenticates via device code flow
  2. Creates three demo users (idempotent): demo_sponsor, demo_approver,
     demo_manager — the human identities the demo pre-fills in setup
  3. Sets the sponsor's manager to demo_manager (required for the Lifecycle
     Workflow offboarding step, which transfers sponsorship to the manager)
  4. Optionally creates a security group and adds all three (handy to exclude
     from a Conditional Access policy so they can sign in from any device)
  5. Outputs the VITE_* values to paste into web/.env.local

  Resolution in the app is by UPN, so you can freely rename the display names
  later; only the UPNs must match what you put in web/.env.local.

.NOTES
  Requires: Microsoft.Graph PowerShell module
  (Install-Module Microsoft.Graph -Scope CurrentUser)

.EXAMPLE
  ./New-DemoUsers.ps1
  Creates the three users in the signed-in admin's tenant (domain inferred from
  the admin's UPN), with a generated password.

.EXAMPLE
  ./New-DemoUsers.ps1 -DomainName contoso.onmicrosoft.com -CreateGroup
  Uses an explicit domain and also creates the "Demo Users" group.
#>

param(
    [string]$DomainName,
    [string]$Password,
    [string]$SponsorName = "demo_sponsor",
    [string]$ApproverName = "demo_approver",
    [string]$ManagerName = "demo_manager",
    [switch]$CreateGroup,
    [string]$GroupName = "Demo Users",
    [switch]$ForcePasswordChange,
    [switch]$ResetPassword,
    [switch]$SkipEnvFile,
    [string]$EnvPath = (Join-Path $PSScriptRoot "..\web\.env.local")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Upserts KEY=value pairs into an env file, preserving any other lines.
function Set-EnvValues {
    param(
        [string]$Path,
        [System.Collections.Specialized.OrderedDictionary]$Values
    )
    $lines = @()
    if (Test-Path $Path) { $lines = @(Get-Content -LiteralPath $Path) }

    foreach ($key in $Values.Keys) {
        $entry = "$key=$($Values[$key])"
        $matched = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*$([regex]::Escape($key))\s*=") {
                $lines[$i] = $entry
                $matched = $true
                break
            }
        }
        if (-not $matched) { $lines += $entry }
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8
}

# --- 0. Ensure required Microsoft.Graph modules are loaded ---
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups"
)
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -Name $mod)) {
        if (Get-Module -ListAvailable -Name $mod) {
            Write-Host "Importing module: $mod ..." -ForegroundColor DarkGray
            Import-Module $mod -ErrorAction Stop
        } else {
            Write-Host "Module '$mod' is not installed." -ForegroundColor Red
            Write-Host "Install it with: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
            exit 1
        }
    }
}

# --- 1. Connect via device code flow ---
Write-Host "`n=== GOV4Agents Demo Users Script ===" -ForegroundColor Cyan
Write-Host "Connecting to Microsoft Graph via device code flow...`n" -ForegroundColor Yellow

$scopes = @("User.ReadWrite.All")
if ($CreateGroup) { $scopes += "Group.ReadWrite.All" }
Connect-MgGraph -Scopes $scopes -UseDeviceCode

$context = Get-MgContext
$tenantId = $context.TenantId
Write-Host "Connected to tenant: $tenantId" -ForegroundColor Green

# --- 2. Resolve the domain for the new UPNs ---
if (-not $DomainName) {
    # Default to the signed-in admin's domain (no extra scope required).
    if ($context.Account -and $context.Account.Contains("@")) {
        $DomainName = $context.Account.Split("@")[-1]
    } else {
        Write-Host "Could not infer domain from the signed-in account." -ForegroundColor Red
        Write-Host "Re-run with -DomainName <tenant>.onmicrosoft.com" -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        exit 1
    }
}
Write-Host "Using domain: $DomainName" -ForegroundColor Green

# --- 3. Generate a password if none was provided ---
function New-DemoPassword {
    $upper = -join ((65..90)   | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    $lower = -join ((97..122)  | Get-Random -Count 6 | ForEach-Object { [char]$_ })
    $digit = -join ((48..57)   | Get-Random -Count 3 | ForEach-Object { [char]$_ })
    $sym   = -join (('!','@','#','$','%','^','&','*') | Get-Random -Count 2)
    return ($upper + $lower + $digit + $sym)
}
if (-not $Password) { $Password = New-DemoPassword }

# --- 4. Define the personas to create ---
$personas = @(
    [ordered]@{ Key = "sponsor";  DisplayName = $SponsorName;  Mail = "demo_sponsor";  Role = "Sponsor — requests access on behalf of the agent" },
    [ordered]@{ Key = "approver"; DisplayName = $ApproverName; Mail = "demo_approver"; Role = "Approver — approves access package requests" },
    [ordered]@{ Key = "manager";  DisplayName = $ManagerName;  Mail = "demo_manager";  Role = "Manager — receives transferred sponsorship at offboarding" }
)

$created = @{}

foreach ($p in $personas) {
    $upn = "$($p.Mail)@$DomainName"
    $existing = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($existing) {
        Write-Host "`nUser '$upn' already exists (idempotent)." -ForegroundColor Yellow
        if ($existing.DisplayName -ne $p.DisplayName) {
            Update-MgUser -UserId $existing.Id -DisplayName $p.DisplayName
            Write-Host "  Display name updated to '$($p.DisplayName)'." -ForegroundColor Green
        }
        if ($ResetPassword) {
            Update-MgUser -UserId $existing.Id -PasswordProfile @{
                Password                      = $Password
                ForceChangePasswordNextSignIn = [bool]$ForcePasswordChange
            }
            Write-Host "  Password reset." -ForegroundColor Green
        }
        $created[$p.Key] = $existing
    } else {
        Write-Host "`nCreating user: $($p.DisplayName) <$upn> ..." -ForegroundColor Yellow
        $user = New-MgUser -BodyParameter @{
            AccountEnabled    = $true
            DisplayName       = $p.DisplayName
            MailNickname      = $p.Mail
            UserPrincipalName = $upn
            PasswordProfile   = @{
                Password                      = $Password
                ForceChangePasswordNextSignIn = [bool]$ForcePasswordChange
            }
        }
        Write-Host "  Created: $($user.Id)" -ForegroundColor Green
        $created[$p.Key] = $user
    }
}

# --- 5. Set the sponsor's manager = demo_manager ---
Write-Host "`nSetting manager: $($created.sponsor.DisplayName) -> $($created.manager.DisplayName) ..." -ForegroundColor Yellow
try {
    Set-MgUserManagerByRef -UserId $created.sponsor.Id -BodyParameter @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($created.manager.Id)"
    }
    Write-Host "  Manager set." -ForegroundColor Green
} catch {
    Write-Host "  Warning: could not set manager: $_" -ForegroundColor Yellow
}

# --- 6. Optional: security group for CA exclusion ---
$group = $null
if ($CreateGroup) {
    Write-Host "`nEnsuring security group '$GroupName' ..." -ForegroundColor Yellow
    $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -All | Select-Object -First 1
    if (-not $group) {
        $group = New-MgGroup -DisplayName $GroupName -MailEnabled:$false `
            -MailNickname ("demo-users-" + [Guid]::NewGuid().ToString("N").Substring(0, 8)) `
            -SecurityEnabled:$true
        Write-Host "  Group created: $($group.Id)" -ForegroundColor Green
    } else {
        Write-Host "  Group exists: $($group.Id)" -ForegroundColor Green
    }

    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
    foreach ($p in $personas) {
        $u = $created[$p.Key]
        $isMember = $members | Where-Object { $_.Id -eq $u.Id }
        if (-not $isMember) {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $u.Id
            Write-Host "  Added $($u.DisplayName) to group." -ForegroundColor Green
        }
    }
}

# --- 7. Output configuration ---
$sponsorUpn  = $created.sponsor.UserPrincipalName
$approverUpn = $created.approver.UserPrincipalName
$managerUpn  = $created.manager.UserPrincipalName

$envValues = [ordered]@{
    "VITE_SPONSOR_UPN"   = $sponsorUpn
    "VITE_SPONSOR_NAME"  = $created.sponsor.DisplayName
    "VITE_APPROVER_UPN"  = $approverUpn
    "VITE_APPROVER_NAME" = $created.approver.DisplayName
    "VITE_MANAGER_UPN"   = $managerUpn
    "VITE_MANAGER_NAME"  = $created.manager.DisplayName
}

Write-Host "`n`n====================================================" -ForegroundColor Cyan
Write-Host "  DEMO USERS COMPLETE" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "--- web/.env.local (persona overrides) ---" -ForegroundColor Yellow
foreach ($k in $envValues.Keys) { Write-Host "$k=$($envValues[$k])" }

if (-not $SkipEnvFile) {
    try {
        Set-EnvValues -Path $EnvPath -Values $envValues
        $resolved = (Resolve-Path -LiteralPath $EnvPath).Path
        Write-Host "`nUpdated env file: $resolved" -ForegroundColor Green
    } catch {
        Write-Host "`nWarning: could not write env file '$EnvPath': $_" -ForegroundColor Yellow
        Write-Host "Copy the values above into web/.env.local manually." -ForegroundColor DarkYellow
    }
} else {
    Write-Host "`n(-SkipEnvFile set: copy the values above into web/.env.local manually.)" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "--- Sign-in password (all three users) ---" -ForegroundColor Yellow
Write-Host $Password
try {
    Set-Clipboard -Value $Password
    Write-Host "(Password copied to clipboard.)" -ForegroundColor Green
} catch {
    Write-Host "(Could not copy to clipboard — copy it manually above.)" -ForegroundColor DarkYellow
}
if ($CreateGroup -and $group) {
    Write-Host ""
    Write-Host "--- Group ---" -ForegroundColor Yellow
    Write-Host "Name : $GroupName"
    Write-Host "Id   : $($group.Id)"
}
Write-Host ""
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NOTES:" -ForegroundColor Yellow
Write-Host "  * No M365 licenses are assigned — the demo is portal-driven and the" -ForegroundColor DarkGray
Write-Host "    personas do not need mailboxes." -ForegroundColor DarkGray
Write-Host "  * On first sign-in the users may be prompted to register MFA /" -ForegroundColor DarkGray
Write-Host "    Authenticator — leave that to the demo participants." -ForegroundColor DarkGray
Write-Host "  * If a Conditional Access policy blocks sign-in (e.g. error 53003)," -ForegroundColor DarkGray
Write-Host "    exclude the group above from the blocking policy (run with" -ForegroundColor DarkGray
Write-Host "    -CreateGroup if you skipped it)." -ForegroundColor DarkGray
Write-Host ""

Disconnect-MgGraph | Out-Null
Write-Host "Done! Disconnected from Graph.`n" -ForegroundColor Green
