using namespace System.Net
function Invoke-SharePointUpload {
    <#
    .SYNOPSIS
        POST /api/SharePointUpload
    .DESCRIPTION
        Body (JSON):
          attachment_name    required  File name to save as
          attachment         required  Base64 file content (e.g. Graph fileAttachment contentBytes)
          folder_path        optional  Folder inside the library. Defaults to $env:SPDefaultFolder
          site_id            optional  Graph site id. Defaults to $env:SPSiteId (uses the site's default library)
          drive_id           optional  Target a specific library instead of the site's default one
          conflict_behavior  optional  replace (default) | rename | fail
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $Body = $Request.Body
    if ($Body -is [string]) {
        try { $Body = $Body | ConvertFrom-Json } catch { $Body = $null }
    }

    $StatusCode = [HttpStatusCode]::OK
    try {
        if (-not $Body -or -not $Body.attachment -or -not $Body.attachment_name) {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw 'attachment and attachment_name are required'
        }

        try {
            $FileBytes = [Convert]::FromBase64String($Body.attachment)
        } catch {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw 'attachment is not valid base64'
        }

        if ($Body.drive_id) {
            $DriveRoot = 'drives/{0}' -f $Body.drive_id
        } else {
            $SiteId = if ($Body.site_id) { $Body.site_id } else { $env:SPSiteId }
            if (-not $SiteId) {
                $StatusCode = [HttpStatusCode]::BadRequest
                throw 'No site_id or drive_id supplied and SPSiteId app setting is not set'
            }
            $DriveRoot = 'sites/{0}/drive' -f $SiteId
        }

        $Conflict = if ($Body.conflict_behavior) { $Body.conflict_behavior } else { 'replace' }
        if ($Conflict -notin 'replace', 'rename', 'fail') {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw "conflict_behavior must be replace, rename or fail"
        }

        $Upload = @{
            DriveRoot        = $DriveRoot
            FolderPath       = if ($null -ne $Body.folder_path) { $Body.folder_path } else { $env:SPDefaultFolder }
            FileName         = $Body.attachment_name
            FileBytes        = $FileBytes
            ConflictBehavior = $Conflict
        }

        $Item = Invoke-GraphUpload @Upload

        $ResponseBody = [ordered]@{
            id      = $Item.id
            name    = $Item.name
            size    = $Item.size
            webUrl  = $Item.webUrl
            driveId = $Item.parentReference.driveId
        }
    } catch {
        if ($StatusCode -eq [HttpStatusCode]::OK) { $StatusCode = [HttpStatusCode]::BadGateway }
        Write-Warning "SharePointUpload failed: $($_.Exception.Message)"
        $ResponseBody = @{ error = $_.Exception.Message }
    }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = $StatusCode
            Body       = $ResponseBody
        })
}
