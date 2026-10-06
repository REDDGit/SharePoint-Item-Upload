using namespace System.Net

function Receive-HttpTrigger {
    Param($Request, $TriggerMetadata)
    Set-Location (Get-Item $PSScriptRoot).Parent.Parent.FullName
    $FunctionName = 'Invoke-{0}' -f $Request.Params.Endpoint

    # Only route to functions exported by the SharePointUpload module, not any Invoke-* command on the box
    $Allowed = (Get-Module SharePointUpload).ExportedFunctions.Keys
    if ($FunctionName -notin $Allowed) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::NotFound
                Body       = @{ error = "Unknown endpoint '$($Request.Params.Endpoint)'" }
            })
        return
    }

    $HttpTrigger = @{
        Request         = $Request
        TriggerMetadata = $TriggerMetadata
    }

    & $FunctionName @HttpTrigger
}

Export-ModuleMember -Function @('Receive-HttpTrigger')
