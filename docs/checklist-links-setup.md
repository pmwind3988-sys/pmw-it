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

## 7. Emails when a checklist is signed (optional)

When an employee signs, the server can email a group IT chooses (Shared
checklists → **Email settings**). The email is sent from a mailbox, by the same
server app, so the app needs permission to send as that ONE mailbox.

1. Pick or create the sending mailbox, e.g. a shared mailbox
   `it-portal@pmwgroup.com` (a shared mailbox needs no licence).
2. On the server app → **API permissions → Add a permission → Microsoft Graph
   → Application permissions → `Mail.Send`** → **Grant admin consent**.
3. Limit it to that mailbox. On its own, `Mail.Send` lets the app send as
   ANYONE in the company. In Exchange Online PowerShell (as an Exchange admin):

   ```powershell
   Connect-ExchangeOnline
   New-DistributionGroup -Name "IT Portal mail senders" -Type Security -PrimarySmtpAddress it-portal-senders@pmwgroup.com
   Add-DistributionGroupMember -Identity it-portal-senders@pmwgroup.com -Member it-portal@pmwgroup.com
   New-ApplicationAccessPolicy -AppId <client-id> -PolicyScopeGroupId it-portal-senders@pmwgroup.com -AccessRight RestrictAccess -Description "PMW IT Portal server: send as it-portal only"
   Test-ApplicationAccessPolicy -AppId <client-id> -Identity it-portal@pmwgroup.com   # AccessCheckResult: Granted
   Test-ApplicationAccessPolicy -AppId <client-id> -Identity someone.else@pmwgroup.com # AccessCheckResult: Denied
   ```

   The policy can take up to an hour to apply.
4. In Vercel add, then redeploy:

   | Name | Value |
   |---|---|
   | `CHECKLIST_MAIL_FROM` | the sending mailbox, e.g. `it-portal@pmwgroup.com` |
   | `PORTAL_URL` | the portal's address, e.g. `https://it.pmwgroup.com` (optional — falls back to `VITE_REDIRECT_URI`) |

   `VITE_CLIENT_ID` (already set) is read by the server too: only sign-in
   tokens issued to the portal may ask for a resend.
5. In the portal: Shared checklists → **Email settings** → add the
   distribution group, tick *Send an email for each new signing*, **Save**,
   then **Send a test email**.

Without `CHECKLIST_MAIL_FROM` everything else works as before and nothing is
emailed. A send that fails never stops a checklist being signed; the link's
panel shows "Email not sent" with the reason, and **Try again**.

Only checklists signed after emails are switched on are sent automatically.
Older ones show "Not emailed" in their panel, with **Send email**.

## Trying it without any of this

```bash
npm run dev:links
```

Opens on port 5174 with an in-memory SharePoint and three demo links:
`/c/DemoInLink`, `/c/DemoOutLnk`, `/c/DemoReqLnk`. Nothing is saved; a restart
resets them. Emails are on in the demo and are only printed to the terminal as
`[dev mail] "<subject>" → <recipients>`.
