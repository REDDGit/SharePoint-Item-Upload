using namespace System.Net
function Invoke-SharePointUpload {
    <#
    .SYNOPSIS
        POST /api/SharePointUpload
    .DESCRIPTION
        Body (JSON):
          attachment_name    required  File name to save as
          attachment         required  Base64 file content (e.g. Graph fileAttachment contentBytes)
          site_id            required* Graph site id (uploads to the site's default library)
          drive_id           required* Specific library id. *Supply site_id or drive_id, not both
          folder_path        required  Folder inside the library. Use "/" for the library root
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

        # No defaults: every request must say exactly where the file goes
        if ([bool]$Body.site_id -eq [bool]$Body.drive_id) {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw 'Supply either site_id or drive_id (exactly one)'
        }
        $DriveRoot = if ($Body.drive_id) { 'drives/{0}' -f $Body.drive_id } else { 'sites/{0}/drive' -f $Body.site_id }

        if ([string]::IsNullOrWhiteSpace($Body.folder_path)) {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw 'folder_path is required. Use "/" to upload to the library root'
        }

        $Conflict = if ($Body.conflict_behavior) { $Body.conflict_behavior } else { 'replace' }
        if ($Conflict -notin 'replace', 'rename', 'fail') {
            $StatusCode = [HttpStatusCode]::BadRequest
            throw "conflict_behavior must be replace, rename or fail"
        }

        $Upload = @{
            DriveRoot        = $DriveRoot
            FolderPath       = $Body.folder_path
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
