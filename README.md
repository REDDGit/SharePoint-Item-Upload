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
  "site_id": "reddgroup.sharepoint.com,fc39acf7-ba08-4b85-be3c-861ab5002ca1,202e4234-2829-4927-bfea-836e7474add2",
  "folder_path": "Backup Radar Reports",
  "conflict_behavior": "replace"
}
```

| Field | Required | Notes |
|---|---|---|
| `attachment_name` | yes | Saved file name. Characters SharePoint rejects (`" * : < > ? / \ |`) become `_` |
| `attachment` | yes | Base64 content |
| `site_id` | one of | Graph site id. Uploads to the site's default library (Shared Documents) |
| `drive_id` | one of | A specific library. Send `site_id` or `drive_id`, not both |
| `folder_path` | yes | Folder inside the library. Use `/` for the library root. Nested paths are fine (`Backup Radar Reports/2026`); missing folders get created |
| `conflict_behavior` | no | `replace` (default), `rename` or `fail` |

Response (200):

```json
{ "id": "...", "name": "...", "size": 516559, "webUrl": "https://reddgroup.sharepoint.com/...", "driveId": "..." }
```

There are no default site or folder settings. A request that doesn't say where the file goes is rejected rather than landing somewhere unintended.

A bad request returns 400 and a Graph failure returns 502, both with `{ "error": "..." }`.

Files up to 4 MB go up in a single PUT. Anything bigger goes through an upload session in 10 MiB chunks.

## Auth

The app uses its system-assigned managed identity to get Graph tokens, so there are no secrets to store or rotate. The identity holds `Sites.Selected`, so it can only write to sites that have been explicitly granted. Any other `site_id` gets a 403 from Graph.

For local testing without a managed identity, set `GraphTenantId`, `GraphClientId` and `GraphClientSecret` in `local.settings.json` for an app registration that has the same access.

## Setup

1. Deploy `Deployment/AzureDeployment.json`.
2. Connect GitHub deployment in the Function App's Deployment Center, the same way as CWDocumentUpload.
3. Run `Deployment/Grant-GraphPermissions.ps1 -FunctionAppName <app name> -SiteId <site id>` for each site the app should write to (needs the Microsoft.Graph PowerShell module and an admin).
4. Restart the Function App.
5. Copy the function key: Functions > HttpTrigger > Function Keys.
