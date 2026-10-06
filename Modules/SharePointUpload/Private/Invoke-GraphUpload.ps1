function Invoke-GraphUpload {
    <#
    .SYNOPSIS
        Uploads raw bytes to a drive path. Simple PUT up to 4 MB, upload session above that.
        Returns the Graph driveItem.
    #>
    Param(
        [Parameter(Mandatory)][string]$DriveRoot,
        [string]$FolderPath,
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][byte[]]$FileBytes,
        [ValidateSet('replace', 'rename', 'fail')][string]$ConflictBehavior = 'replace'
    )

    # SharePoint won't accept these characters in a file name
    $SafeName = ($FileName -replace '["*:<>?/\\|]', '_').Trim().TrimEnd('.')

    $Segments = @(($FolderPath -split '[/\\]') | Where-Object { $_.Trim() }) + $SafeName
    $ItemPath = ($Segments | ForEach-Object { [Uri]::EscapeDataString($_.Trim()) }) -join '/'
    $ItemUri = 'https://graph.microsoft.com/v1.0/{0}/root:/{1}' -f $DriveRoot, $ItemPath

    $Headers = @{ Authorization = 'Bearer {0}' -f (Get-GraphToken) }
    $SimpleLimit = 4MB

    if ($FileBytes.Length -le $SimpleLimit) {
        $Put = @{
            Uri                = '{0}:/content?@microsoft.graph.conflictBehavior={1}' -f $ItemUri, $ConflictBehavior
            Method             = 'Put'
            Headers            = $Headers
            Body               = $FileBytes
            ContentType        = 'application/octet-stream'
            SkipHttpErrorCheck = $true
            StatusCodeVariable = 'StatusCode'
        }
        $Result = Invoke-RestMethod @Put
        if ($StatusCode -notin 200, 201) {
            throw ('Graph upload failed ({0}): {1}' -f $StatusCode, ($Result | ConvertTo-Json -Depth 5 -Compress))
        }
        return $Result
    }

    # Large file: upload session
    $Session = @{
        Uri                = '{0}:/createUploadSession' -f $ItemUri
        Method             = 'Post'
        Headers            = $Headers
        Body               = (@{ item = @{ '@microsoft.graph.conflictBehavior' = $ConflictBehavior } } | ConvertTo-Json)
        ContentType        = 'application/json'
        SkipHttpErrorCheck = $true
        StatusCodeVariable = 'StatusCode'
    }
    $UploadSession = Invoke-RestMethod @Session
    if ($StatusCode -ne 200 -or -not $UploadSession.uploadUrl) {
        throw ('Could not create upload session ({0}): {1}' -f $StatusCode, ($UploadSession | ConvertTo-Json -Depth 5 -Compress))
    }

    $ChunkSize = 320KB * 32   # 10 MiB, must be a multiple of 320 KiB
    $Total = $FileBytes.Length
    for ($Offset = 0; $Offset -lt $Total; $Offset += $ChunkSize) {
        $Length = [Math]::Min($ChunkSize, $Total - $Offset)
        $Chunk = [byte[]]::new($Length)
        [Array]::Copy($FileBytes, $Offset, $Chunk, 0, $Length)

        # No Authorization header on uploadUrl - it's pre-authenticated
        $Part = @{
            Uri                  = $UploadSession.uploadUrl
            Method               = 'Put'
            Headers              = @{ 'Content-Range' = 'bytes {0}-{1}/{2}' -f $Offset, ($Offset + $Length - 1), $Total }
            Body                 = $Chunk
            ContentType          = 'application/octet-stream'
            SkipHeaderValidation = $true
            SkipHttpErrorCheck   = $true
            StatusCodeVariable   = 'StatusCode'
        }
        $Result = Invoke-RestMethod @Part
        if ($StatusCode -notin 200, 201, 202) {
            Invoke-RestMethod -Uri $UploadSession.uploadUrl -Method Delete -SkipHttpErrorCheck | Out-Null
            throw ('Chunk upload failed at byte {0} ({1}): {2}' -f $Offset, $StatusCode, ($Result | ConvertTo-Json -Depth 5 -Compress))
        }
    }

    return $Result
}
