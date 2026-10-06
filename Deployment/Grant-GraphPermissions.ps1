<#
.SYNOPSIS
    Gives the Function App's managed identity access to SharePoint through Microsoft Graph.

    -SiteId    Assigns Sites.Selected and grants 'write' on that one site. Run once per site.
    -AllSites  Assigns Sites.ReadWrite.All: read/write on every site in the tenant, no per-site grants.

    Needs an admin who can consent AppRoleAssignment.ReadWrite.All (and Sites.FullControl.All for -SiteId).
    Restart the Function App afterwards so it picks up a fresh token.

.EXAMPLE
    ./Grant-GraphPermissions.ps1 -FunctionAppName sharepoint-item-uploadabcde -SiteId 'reddgroup.sharepoint.com,fc39acf7-ba08-4b85-be3c-861ab5002ca1,202e4234-2829-4927-bfea-836e7474add2'

    Get a site id with: GET https://graph.microsoft.com/v1.0/sites/reddgroup.sharepoint.com:/sites/<site>

.EXAMPLE
    ./Grant-GraphPermissions.ps1 -FunctionAppName sharepoint-item-uploadabcde -AllSites
#>
[CmdletBinding(DefaultParameterSetName = 'Site')]
param(
    [Parameter(Mandatory)][string]$FunctionAppName,
    [Parameter(Mandatory, ParameterSetName = 'Site')][string]$SiteId,
    [Parameter(ParameterSetName = 'Site')][ValidateSet('read', 'write')][string]$Role = 'write',
    [Parameter(Mandatory, ParameterSetName = 'All')][switch]$AllSites
)

$Scopes = @('Application.Read.All', 'AppRoleAssignment.ReadWrite.All')
if (-not $AllSites) { $Scopes += 'Sites.FullControl.All' }
Connect-MgGraph -Scopes $Scopes -NoWelcome

$Identity = Get-MgServicePrincipal -Filter "displayName eq '$FunctionAppName' and servicePrincipalType eq 'ManagedIdentity'"
if (-not $Identity) { throw "No managed identity found for '$FunctionAppName'. Is the system-assigned identity turned on?" }

$GraphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$Permission = if ($AllSites) { 'Sites.ReadWrite.All' } else { 'Sites.Selected' }
$AppRole = $GraphSp.AppRoles | Where-Object { $_.Value -eq $Permission -and $_.AllowedMemberTypes -contains 'Application' }

$Existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $Identity.Id | Where-Object { $_.AppRoleId -eq $AppRole.Id }
if ($Existing) {
    Write-Host "$Permission already assigned"
} else {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $Identity.Id -PrincipalId $Identity.Id -ResourceId $GraphSp.Id -AppRoleId $AppRole.Id | Out-Null
    Write-Host "Assigned $Permission to $($Identity.DisplayName)"
}

if (-not $AllSites) {
    $SiteGrant = @{
        roles               = @($Role)
        grantedToIdentities = @(@{ application = @{ id = $Identity.AppId; displayName = $Identity.DisplayName } })
    } | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions" -Body $SiteGrant -ContentType 'application/json' | Out-Null
    Write-Host "Granted '$Role' on $SiteId to $($Identity.DisplayName) ($($Identity.AppId))"
}

Write-Host 'Restart the Function App so it picks up a fresh token with the new role.'
