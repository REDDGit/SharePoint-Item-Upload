@{
    RootModule        = '.\SharePointUpload.psm1'
    ModuleVersion     = '1.0'
    GUID              = '3f6b1c2e-8d4a-4e7b-9a51-2c7d0e9f4b18'
    Author            = 'REDD'
    Description       = 'Uploads base64 file content to a SharePoint document library via Microsoft Graph'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Invoke-SharePointUpload')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
