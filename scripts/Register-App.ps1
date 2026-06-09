<#
.SYNOPSIS
  Registers the GOV4Agents app registration in a target tenant using device code flow.
  
.DESCRIPTION
  This script:
  1. Authenticates via device code flow (no app needed beforehand)
  2. Creates a single-tenant app registration with all required API permissions
  3. Exposes an API scope (access_as_user) for the SPA
  4. Configures the SPA redirect URI (http://localhost:5173)
  5. Creates a client secret
  6. Grants admin consent for all permissions
  7. Outputs the values needed for .env files

.NOTES
  Requires: Microsoft.Graph PowerShell module (Install-Module Microsoft.Graph -Scope CurrentUser)
#>

param(
    [string]$AppDisplayName = "GOV4Agents Demo",
    [string]$RedirectUri = "http://localhost:5173",
    [switch]$NewSecret
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- 0. Ensure required Microsoft.Graph modules are loaded ---
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")
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
Write-Host "`n=== GOV4Agents App Registration Script ===" -ForegroundColor Cyan
Write-Host "Connecting to Microsoft Graph via device code flow...`n" -ForegroundColor Yellow

Connect-MgGraph -Scopes @(
    "Application.ReadWrite.All",
    "DelegatedPermissionGrant.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All"
) -UseDeviceCode

$context = Get-MgContext
$tenantId = $context.TenantId
Write-Host "Connected to tenant: $tenantId" -ForegroundColor Green

# --- 2. Define required permissions ---
# Microsoft Graph App ID (constant across tenants)
$graphAppId = "00000003-0000-0000-c000-000000000000"

# Delegated permissions (scope type)
$delegatedPermissions = @(
    "User.Read",
    "Directory.ReadWrite.All",
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "EntitlementManagement.ReadWrite.All",
    "LifecycleWorkflows.ReadWrite.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Policy.Read.All",
    "CustomSecAttributeDefinition.ReadWrite.All",
    "CustomSecAttributeAssignment.ReadWrite.All",
    "AuditLog.Read.All",
    # --- Agent Identity preview scopes (required for Steps 1 & 2) ---
    "AgentIdentity.ReadWrite.All",
    "AgentIdentityBlueprint.Create",
    "AgentIdentityBlueprint.UpdateAuthProperties.All",
    "AgentIdentityBlueprint.AddRemoveCreds.All",
    "AgentIdentityBlueprintPrincipal.Create"
)

# Application permissions (role type) - for client credentials flow
$applicationPermissions = @(
    "Application.ReadWrite.All",
    "Directory.ReadWrite.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Policy.Read.All",
    "CustomSecAttributeDefinition.ReadWrite.All",
    "CustomSecAttributeAssignment.ReadWrite.All"
)

# --- 3. Look up permission IDs from the Graph service principal ---
Write-Host "`nLooking up permission IDs from Microsoft Graph..." -ForegroundColor Yellow
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Select "id,appId,oauth2PermissionScopes,appRoles"

$delegatedIds = @()
$grantedScopes = @()
foreach ($perm in $delegatedPermissions) {
    $found = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $perm }
    if ($found) {
        $delegatedIds += @{ Id = $found.Id; Type = "Scope" }
        $grantedScopes += $perm
        Write-Host "  [Delegated] $perm => $($found.Id)" -ForegroundColor DarkGray
    } else {
        Write-Host "  [Delegated] $perm => NOT FOUND in this tenant's Graph SP (skipped)" -ForegroundColor DarkYellow
        # Preview scopes absent from the SP manifest cannot be consented; the
        # Agent ID preview must be enabled in the tenant for these to appear.
    }
}

$applicationIds = @()
foreach ($perm in $applicationPermissions) {
    $found = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
    if ($found) {
        $applicationIds += @{ Id = $found.Id; Type = "Role" }
        Write-Host "  [Application] $perm => $($found.Id)" -ForegroundColor DarkGray
    } else {
        Write-Host "  [Application] $perm => NOT FOUND" -ForegroundColor DarkYellow
    }
}

# --- 4. Create or update the app registration (idempotent) ---
$requiredResourceAccess = @(
    @{
        ResourceAppId  = $graphAppId
        ResourceAccess = ($delegatedIds + $applicationIds)
    }
)

$existingApp = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -All | Select-Object -First 1

if ($existingApp) {
    Write-Host "`nApp '$AppDisplayName' already exists - updating (idempotent)..." -ForegroundColor Yellow
    $appExisted = $true
    $appObjectId = $existingApp.Id
    $appId = $existingApp.AppId

    # Preserve the existing access_as_user scope id; only add it if missing
    # (regenerating the id would invalidate the SPA's existing consent).
    $existingScopes = @()
    if ($existingApp.Api -and $existingApp.Api.Oauth2PermissionScopes) {
        $existingScopes = $existingApp.Api.Oauth2PermissionScopes
    }
    $hasAccessScope = $existingScopes | Where-Object { $_.Value -eq 'access_as_user' }

    # Merge SPA redirect URIs (keep any existing ones)
    $spaRedirects = @($RedirectUri)
    if ($existingApp.Spa -and $existingApp.Spa.RedirectUris) {
        $spaRedirects = @($existingApp.Spa.RedirectUris + $RedirectUri | Select-Object -Unique)
    }

    $updateParams = @{
        ApplicationId          = $appObjectId
        RequiredResourceAccess = $requiredResourceAccess
        Spa                    = @{ RedirectUris = $spaRedirects }
    }
    if (-not $hasAccessScope) {
        $updateParams.Api = @{
            RequestedAccessTokenVersion = 2
            Oauth2PermissionScopes      = @(
                @{
                    Id                      = [Guid]::NewGuid().ToString()
                    AdminConsentDisplayName = "Access GOV4Agents API as user"
                    AdminConsentDescription = "Allows the SPA to call the GOV4Agents backend on behalf of the signed-in user"
                    UserConsentDisplayName  = "Access GOV4Agents API"
                    UserConsentDescription  = "Allows access to the GOV4Agents API on your behalf"
                    Value                   = "access_as_user"
                    Type                    = "User"
                    IsEnabled               = $true
                }
            )
        }
    }
    Update-MgApplication @updateParams

    if (-not $existingApp.IdentifierUris -or $existingApp.IdentifierUris.Count -eq 0) {
        Update-MgApplication -ApplicationId $appObjectId -IdentifierUris @("api://$appId")
    }
    Write-Host "  App updated: $appId (object: $appObjectId)" -ForegroundColor Green
} else {
    Write-Host "`nCreating app registration: $AppDisplayName ..." -ForegroundColor Yellow
    $appExisted = $false

    $appBody = @{
        DisplayName            = $AppDisplayName
        SignInAudience         = "AzureADMyOrg"
        RequiredResourceAccess = $requiredResourceAccess
        Spa                    = @{ RedirectUris = @($RedirectUri) }
        Api                    = @{
            RequestedAccessTokenVersion = 2
            Oauth2PermissionScopes      = @(
                @{
                    Id                      = [Guid]::NewGuid().ToString()
                    AdminConsentDisplayName = "Access GOV4Agents API as user"
                    AdminConsentDescription = "Allows the SPA to call the GOV4Agents backend on behalf of the signed-in user"
                    UserConsentDisplayName  = "Access GOV4Agents API"
                    UserConsentDescription  = "Allows access to the GOV4Agents API on your behalf"
                    Value                   = "access_as_user"
                    Type                    = "User"
                    IsEnabled               = $true
                }
            )
        }
        Web                    = @{
            ImplicitGrantSettings = @{
                EnableAccessTokenIssuance = $false
                EnableIdTokenIssuance     = $false
            }
        }
    }

    $app = New-MgApplication -BodyParameter $appBody
    $appId = $app.AppId
    $appObjectId = $app.Id
    Write-Host "  App created: $appId (object: $appObjectId)" -ForegroundColor Green

    Write-Host "Setting Application ID URI..." -ForegroundColor Yellow
    Update-MgApplication -ApplicationId $appObjectId -IdentifierUris @("api://$appId")
}

# --- 5. Client secret (new app, or when -NewSecret is passed) ---
$clientSecret = $null
if (-not $appExisted -or $NewSecret) {
    Write-Host "Creating client secret..." -ForegroundColor Yellow
    $secret = Add-MgApplicationPassword -ApplicationId $appObjectId -PasswordCredential @{
        DisplayName = "GOV4Agents demo secret"
        EndDateTime = (Get-Date).AddMonths(12)
    }
    $clientSecret = $secret.SecretText
} else {
    Write-Host "Reusing existing client secret (pass -NewSecret to rotate)." -ForegroundColor DarkGray
}

# --- 6. Service principal (create only if missing) ---
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" | Select-Object -First 1
if (-not $sp) {
    Write-Host "Creating service principal..." -ForegroundColor Yellow
    $sp = New-MgServicePrincipal -AppId $appId
    Start-Sleep -Seconds 5
}
$spId = $sp.Id

# --- 7. Admin consent for delegated permissions (idempotent) ---
Write-Host "Granting admin consent for delegated permissions..." -ForegroundColor Yellow
$delegatedScopeString = ($grantedScopes -join " ")
try {
    $existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$spId' and resourceId eq '$($graphSp.Id)'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existingGrant) {
        Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id -Scope $delegatedScopeString
        Write-Host "  Updated existing consent grant" -ForegroundColor Green
    } else {
        New-MgOauth2PermissionGrant -BodyParameter @{
            ClientId    = $spId
            ConsentType = "AllPrincipals"
            ResourceId  = $graphSp.Id
            Scope       = $delegatedScopeString
        } | Out-Null
        Write-Host "  Delegated permissions consented" -ForegroundColor Green
    }
} catch {
    Write-Host "  Warning: Could not grant delegated consent: $_" -ForegroundColor Yellow
    Write-Host "  You may need to consent manually in the Azure portal." -ForegroundColor Yellow
}

# --- 8. Admin consent for application permissions (idempotent) ---
Write-Host "Granting admin consent for application permissions..." -ForegroundColor Yellow
$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spId -ErrorAction SilentlyContinue
foreach ($role in $applicationIds) {
    $already = $existingAssignments | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $graphSp.Id }
    if (-not $already) {
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spId -BodyParameter @{
                PrincipalId = $spId
                ResourceId  = $graphSp.Id
                AppRoleId   = $role.Id
            } | Out-Null
        } catch {
            Write-Host "  Warning: Could not assign role $($role.Id): $_" -ForegroundColor Yellow
        }
    }
}
Write-Host "  Application permissions consented" -ForegroundColor Green

# --- 10. Output configuration ---
Write-Host "`n`n====================================================" -ForegroundColor Cyan
Write-Host "  APP REGISTRATION COMPLETE" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "--- server/.env ---" -ForegroundColor Yellow
Write-Host "TENANT_ID=$tenantId"
Write-Host "CLIENT_ID=$appId"
if ($clientSecret) {
    Write-Host "CLIENT_SECRET=<copied to clipboard - paste manually>"
} else {
    Write-Host "CLIENT_SECRET=<unchanged - keep your existing value>"
}
Write-Host "SPA_ORIGIN=$RedirectUri"
Write-Host ""
Write-Host "--- web/.env.local ---" -ForegroundColor Yellow
Write-Host "VITE_CLIENT_ID=$appId"
Write-Host "VITE_AUTHORITY=https://login.microsoftonline.com/$tenantId"
Write-Host ""
if ($clientSecret) {
    Set-Clipboard -Value $clientSecret
    Write-Host "The client secret has been copied to your clipboard." -ForegroundColor Green
    Write-Host "Paste it into server/.env as CLIENT_SECRET value." -ForegroundColor Green
} else {
    Write-Host "Existing app updated - your current CLIENT_SECRET still works." -ForegroundColor Green
    Write-Host "(Re-run with -NewSecret to rotate the secret.)" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Portal link:" -ForegroundColor Yellow
Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$appId" 
Write-Host ""
Write-Host "NOTE: If any Agent Identity preview scope showed 'NOT FOUND in this tenant's" -ForegroundColor DarkYellow
Write-Host "Graph SP', the Agent ID preview is not enabled in this tenant - Steps 1 and 2" -ForegroundColor DarkYellow
Write-Host "will fail until it is. All other scopes are granted and consented." -ForegroundColor DarkYellow
Write-Host ""
Write-Host "----------------------------------------------------" -ForegroundColor Yellow
Write-Host "ACTION REQUIRED: Custom Security Attribute roles" -ForegroundColor Yellow
Write-Host "----------------------------------------------------" -ForegroundColor Yellow
Write-Host "This script did NOT assign directory roles. The user who signs into the"
Write-Host "web UI must hold BOTH of these Entra roles, or Step 5 + the Protect journey"
Write-Host "will return 403 (Global Administrator does NOT include these by default):"
Write-Host "  - Attribute Definition Administrator" -ForegroundColor Cyan
Write-Host "  - Attribute Assignment Administrator" -ForegroundColor Cyan
Write-Host ""
Write-Host "Assign them in Entra (needs a Privileged Role Admin / Global Admin):"
Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles"
Write-Host ""

Disconnect-MgGraph | Out-Null
Write-Host "Done! Disconnected from Graph.`n" -ForegroundColor Green
