<#
.SYNOPSIS
    Gives the Function App's managed identity write access to one SharePoint site only.
    1. Assigns the Graph application permission Sites.Selected to the managed identity
    2. Grants that identity 'write' on the target site

    Run once after deploying. Needs an admin who can consent AppRoleAssignment.ReadWrite.All
    and Sites.FullControl.All (Global Admin / Privileged Role Admin).

.EXAMPLE
    ./Grant-GraphPermissions.ps1 -FunctionAppName spdocumentuploadabcde
#>
param(
    [Parameter(Mandatory)][string]$FunctionAppName,
    [string]$SiteId = 'reddgroup.sharepoint.com,fc39acf7-ba08-4b85-be3c-861ab5002ca1,202e4234-2829-4927-bfea-836e7474add2',
    [ValidateSet('read', 'write')][string]$Role = 'write'
)

Connect-MgGraph -Scopes 'Application.Read.All', 'AppRoleAssignment.ReadWrite.All', 'Sites.FullControl.All' -NoWelcome

$Identity = Get-MgServicePrincipal -Filter "displayName eq '$FunctionAppName' and servicePrincipalType eq 'ManagedIdentity'"
if (-not $Identity) { throw "No managed identity found for '$FunctionAppName'. Is the system-assigned identity turned on?" }

$GraphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$SitesSelected = $GraphSp.AppRoles | Where-Object { $_.Value -eq 'Sites.Selected' -and $_.AllowedMemberTypes -contains 'Application' }

$Existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $Identity.Id | Where-Object { $_.AppRoleId -eq $SitesSelected.Id }
if ($Existing) {
    Write-Host 'Sites.Selected already assigned'
} else {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $Identity.Id -PrincipalId $Identity.Id -ResourceId $GraphSp.Id -AppRoleId $SitesSelected.Id | Out-Null
    Write-Host 'Assigned Sites.Selected'
}

$Permission = @{
    roles               = @($Role)
    grantedToIdentities = @(@{ application = @{ id = $Identity.AppId; displayName = $Identity.DisplayName } })
} | ConvertTo-Json -Depth 5

Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions" -Body $Permission -ContentType 'application/json' | Out-Null
Write-Host "Granted '$Role' on $SiteId to $($Identity.DisplayName) ($($Identity.AppId))"
Write-Host 'Restart the Function App so it picks up a fresh token with the new role.'
