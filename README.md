# SharePoint-Item-Upload

Azure Function (PowerShell) that takes a base64 file, e.g. an Outlook attachment's `contentBytes`, decodes it and uploads it to a SharePoint document library through Microsoft Graph. It exists because Rewst can't send a binary body to Graph itself.

Built on the same structure as CWDocumentUpload.

## Request

`POST https://<app>.azurewebsites.net/api/SharePointUpload`
Header: `x-functions-key: <function key>`

```json
{
  "attachment_name": "Department of Premier and Cabinet - (MIN)-20260812.pdf",
  "attachment": "<contentBytes from the Graph attachment>",
  "folder_path": "Backup Radar Reports",
  "conflict_behavior": "replace"
}
```

| Field | Required | Notes |
|---|---|---|
| `attachment_name` | yes | Saved file name. Characters SharePoint rejects (`" * : < > ? / \ |`) become `_` |
| `attachment` | yes | Base64 content |
| `folder_path` | no | Defaults to the `SPDefaultFolder` app setting. Nested paths are fine (`Backup Radar Reports/2026`); missing folders get created |
| `site_id` | no | Defaults to the `SPSiteId` app setting. Uses the site's default library (Shared Documents) |
| `drive_id` | no | Targets a specific library instead of the site default |
| `conflict_behavior` | no | `replace` (default), `rename` or `fail` |

Response (200):

```json
{ "id": "...", "name": "...", "size": 516559, "webUrl": "https://reddgroup.sharepoint.com/...", "driveId": "..." }
```

A bad request returns 400 and a Graph failure returns 502, both with `{ "error": "..." }`.

Files up to 4 MB go up in a single PUT. Anything bigger goes through an upload session in 10 MiB chunks.

## Auth

The app uses its system-assigned managed identity to get Graph tokens, so there are no secrets to store or rotate. The identity holds `Sites.Selected` and has write on the NOC site only.

For local testing without a managed identity, set `GraphTenantId`, `GraphClientId` and `GraphClientSecret` in `local.settings.json` for an app registration that has the same access.

## Setup

1. Create the repo and deploy `Deployment/AzureDeployment.json`. The defaults already point at the NOC site and `Backup Radar Reports`.
2. Connect GitHub deployment in the Function App's Deployment Center, the same way as CWDocumentUpload.
3. Run `Deployment/Grant-GraphPermissions.ps1 -FunctionAppName <app name>` (needs the Microsoft.Graph PowerShell module and an admin).
4. Restart the Function App.
5. Copy the function key: Functions > HttpTrigger > Function Keys.

## App settings

| Name | Purpose |
|---|---|
| `SPSiteId` | Default Graph site id |
| `SPDefaultFolder` | Default folder in the library |
