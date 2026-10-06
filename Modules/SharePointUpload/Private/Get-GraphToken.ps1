function Get-GraphToken {
    <#
    .SYNOPSIS
        Returns a Graph access token.
        In Azure: uses the Function App's system-assigned managed identity (no secrets).
        Locally: falls back to client credentials if GraphTenantId / GraphClientId / GraphClientSecret are set.
    #>
    if ($script:GraphToken -and $script:GraphTokenExpiry -gt (Get-Date).ToUniversalTime().AddMinutes(5)) {
        return $script:GraphToken
    }

    if ($env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER) {
        $Uri = '{0}?resource={1}&api-version=2019-08-01' -f $env:IDENTITY_ENDPOINT, [Uri]::EscapeDataString('https://graph.microsoft.com')
        $Token = Invoke-RestMethod -Uri $Uri -Method Get -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }
        $script:GraphTokenExpiry = [DateTimeOffset]::FromUnixTimeSeconds([long]$Token.expires_on).UtcDateTime
    } elseif ($env:GraphClientSecret) {
        $TokenRequest = @{
            Uri         = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $env:GraphTenantId
            Method      = 'Post'
            ContentType = 'application/x-www-form-urlencoded'
            Body        = @{
                client_id     = $env:GraphClientId
                client_secret = $env:GraphClientSecret
                scope         = 'https://graph.microsoft.com/.default'
                grant_type    = 'client_credentials'
            }
        }
        $Token = Invoke-RestMethod @TokenRequest
        $script:GraphTokenExpiry = (Get-Date).ToUniversalTime().AddSeconds([int]$Token.expires_in)
    } else {
        throw 'No managed identity available and no GraphClientSecret configured'
    }

    $script:GraphToken = $Token.access_token
    return $script:GraphToken
}
