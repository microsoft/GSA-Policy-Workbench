#Requires -Version 7
#Requires -Modules Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Creates a Microsoft Entra app registration for local SPA dev and writes the env file.

.DESCRIPTION
    Connects to the specified tenant using Microsoft Graph PowerShell (browser sign-in),
    creates a single-tenant SPA app registration with the required delegated permissions,
    optionally grants admin consent, and writes VITE_AAD_CLIENT_ID / VITE_AAD_TENANT to
    the specified env file.

    TenantId is mandatory — the connection is always scoped to the exact tenant you
    specify, so the app registration is always created in the right place.

.PARAMETER Name
    Display name for the app registration.
    Default: "GSA Policy Workbench (local dev)".

.PARAMETER TenantId
    The tenant in which to create the app registration. Mandatory.

.PARAMETER RedirectUri
    SPA redirect URI.
    Default: http://localhost:5173.

.PARAMETER EnvPath
    Path to the env file to write, relative to the current directory.
    Default: .env.local.

.PARAMETER Force
    Overwrite the env file if it already exists.

.PARAMETER SkipAdminConsent
    Skip the admin consent step. A tenant admin will need to grant consent manually.

.EXAMPLE
    ./CreateAppRegistration.ps1 -TenantId xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx

.EXAMPLE
    ./CreateAppRegistration.ps1 `
        -TenantId xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx `
        -Name "GSA Policy Workbench" `
        -EnvPath .env.contoso.local `
        -Force
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]  $Name            = 'GSA Policy Workbench (local dev)',
    [Parameter(Mandatory)]
    [string]  $TenantId,
    [string]  $RedirectUri     = 'http://localhost:5173',
    [string]  $EnvPath         = '.env.local',
    [switch]  $Force,
    [switch]  $SkipAdminConsent
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Well-known Microsoft Graph service principal app ID (same across all tenants).
$GraphAppId = '00000003-0000-0000-c000-000000000000'

# ── 1. Connect to the target tenant ──────────────────────────────────────────

Write-Host "Connecting to tenant $TenantId…"
Connect-MgGraph `
    -TenantId $TenantId `
    -Scopes 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' `
    -NoWelcome

$context = Get-MgContext
if (-not $context) {
    throw 'Connect-MgGraph did not return a context. Authentication may have been cancelled.'
}
# TenantId property name varies by module version; fall back to HomeTenantId.
$connectedTenant = $context.TenantId
if (-not $connectedTenant) { $connectedTenant = $context.HomeTenantId }
if ($connectedTenant -and ($connectedTenant -ne $TenantId)) {
    throw "Connected tenant ($connectedTenant) does not match requested tenant ($TenantId). Aborting."
}
Write-Host "Connected as $($context.Account)"

# ── 2. Resolve delegated permission IDs from the Graph service principal ──────

Write-Host 'Resolving Graph delegated permission IDs…'
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
if (-not $graphSp) {
    throw "Microsoft Graph service principal not found in tenant $TenantId."
}

$networkReadScopeId = ($graphSp.Oauth2PermissionScopes | Where-Object Value -eq 'NetworkAccess.Read.All').Id
$policyReadScopeId  = ($graphSp.Oauth2PermissionScopes | Where-Object Value -eq 'Policy.Read.All').Id

if (-not $networkReadScopeId) { throw 'Could not resolve scope ID for NetworkAccess.Read.All' }
if (-not $policyReadScopeId)  { throw 'Could not resolve scope ID for Policy.Read.All' }

# ── 3. Create the app registration ───────────────────────────────────────────

Write-Host "Creating app registration: $Name"
$requiredAccess = @(
    @{
        ResourceAppId  = $GraphAppId
        ResourceAccess = @(
            @{ Id = $networkReadScopeId; Type = 'Scope' }
            @{ Id = $policyReadScopeId;  Type = 'Scope' }
        )
    }
)

$app = New-MgApplication `
    -DisplayName          $Name `
    -SignInAudience        'AzureADMyOrg' `
    -RequiredResourceAccess $requiredAccess `
    -Spa                  @{ RedirectUris = @($RedirectUri) }

Write-Host "App created — client ID: $($app.AppId)"

# ── 4. Create the service principal (required for admin consent) ──────────────

Write-Host 'Creating service principal…'
$sp = New-MgServicePrincipal -AppId $app.AppId

# ── 5. Grant admin consent ────────────────────────────────────────────────────

if ($SkipAdminConsent) {
    Write-Host 'Skipping admin consent as requested.'
} else {
    Write-Host 'Granting admin consent…'
    try {
        New-MgOauth2PermissionGrant `
            -ClientId    $sp.Id `
            -ConsentType 'AllPrincipals' `
            -ResourceId  $graphSp.Id `
            -Scope       'NetworkAccess.Read.All Policy.Read.All' | Out-Null
        Write-Host 'Admin consent granted.'
    } catch {
        Write-Warning "Admin consent could not be granted automatically. A tenant admin may need to grant consent manually in Entra."
        Write-Warning $_.Exception.Message
    }
}

# ── 6. Write env file ─────────────────────────────────────────────────────────

$fullEnvPath = Join-Path (Get-Location) $EnvPath
if ((Test-Path $fullEnvPath) -and -not $Force) {
    throw "$EnvPath already exists. Re-run with -Force to overwrite."
}

$envContent = @"
# Auto-generated by CreateAppRegistration.ps1
VITE_AAD_CLIENT_ID=$($app.AppId)
VITE_AAD_TENANT=$TenantId
# Optional: override Graph endpoint
# VITE_GRAPH_BASE_URL=https://graph.microsoft.com/beta
"@

Set-Content -Path $fullEnvPath -Value $envContent -Encoding utf8NoBOM
Write-Host "Wrote $EnvPath"

# ── Summary ───────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host 'Done.'
Write-Host "App display name : $Name"
Write-Host "Client ID        : $($app.AppId)"
Write-Host "Tenant           : $TenantId"
Write-Host "Env file         : $EnvPath"
