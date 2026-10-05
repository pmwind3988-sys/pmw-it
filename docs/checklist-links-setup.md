# Shared checklist links — one-time setup

The portal can share an asset checklist as a short link (`/c/<code>`) that an
employee opens **without signing in**, fills, signs and submits. Because that
person has no Microsoft sign-in, the portal saves their checklist under its
**own** identity. That identity has to be created once, by someone with Azure
AD admin rights.

Until it is, links can still be created in the portal, but opening one shows
"This form cannot be opened right now", and the Vercel function log says which
setting is missing.

## 1. Register the server app

Azure portal → **Microsoft Entra ID → App registrations → New registration**

- Name: `PMW IT Portal — server`
- Supported account types: *this organizational directory only*
- Redirect URI: none

Note the **Application (client) ID** and **Directory (tenant) ID**.

## 2. Give it Graph `Sites.Selected`

On the new app → **API permissions → Add a permission → Microsoft Graph →
Application permissions → `Sites.Selected`** → **Grant admin consent**.

`Sites.Selected` on its own grants access to *no* site. Step 3 grants exactly
one.

## 3. Allow it to write to the IT helpdesk site only

Pick one.

**PnP PowerShell** (as a SharePoint admin):

```powershell
Connect-PnPOnline -Url https://pmwgroupcom.sharepoint.com/sites/IThelpdesk -Interactive
Grant-PnPAzureADAppSitePermission -AppId <client-id> -DisplayName "PMW IT Portal — server" -Permissions Write -Site https://pmwgroupcom.sharepoint.com/sites/IThelpdesk
```

**Graph Explorer** (signed in as an admin with `Sites.FullControl.All`):

```
GET https://graph.microsoft.com/v1.0/sites/pmwgroupcom.sharepoint.com:/sites/IThelpdesk?$select=id
```

then, with the `id` it returns:

```
POST https://graph.microsoft.com/v1.0/sites/<site-id>/permissions
{
  "roles": ["write"],
  "grantedToIdentities": [
    { "application": { "id": "<client-id>", "displayName": "PMW IT Portal — server" } }
  ]
}
```

## 4. Create a client secret

On the app → **Certificates & secrets → New client secret**. Copy the
**Value** (shown once). Set a reminder before it expires — when it does, links
stop opening until a new one is put in Vercel.

## 5. Put the settings in Vercel

Project → **Settings → Environment Variables** (Production, and Preview if
wanted):

| Name | Value |
|---|---|
| `CHECKLIST_APP_CLIENT_ID` | the client ID from step 1 |
| `CHECKLIST_APP_CLIENT_SECRET` | the secret value from step 4 |
| `CHECKLIST_TENANT_ID` | the tenant ID (optional — falls back to `VITE_TENANT_ID`) |

`VITE_SHAREPOINT_SITE_URL` is already set and tells the server which site to
use. Redeploy after adding them.

## 6. Check it

1. Sign in to the portal, open **Asset checklist → Share as a link**, create a
   link. The first one also creates the *Asset Checklist Links* list.
2. Open the link in a private window. It should show the form, not an error.
3. Sign and submit. The checklist appears in *Asset Checklist Form* exactly as
   one signed in the portal would, and the link now shows the signed copy.

## Trying it without any of this

```bash
npm run dev:links
```

Opens on port 5174 with an in-memory SharePoint and three demo links:
`/c/DemoInLink`, `/c/DemoOutLnk`, `/c/DemoReqLnk`. Nothing is saved; a restart
resets them.
