# Raise a ticket (IT helpdesk web part)

The "Raise a ticket" form and "My tickets" list on the IT helpdesk SharePoint site
(https://pmwgroupcom.sharepoint.com/sites/IThelpdesk). It replaces the stock list form
and writes to the existing **Tickets** list as the signed-in person — no sign-in or app
permission of its own.

- **The Tickets list is not changed by this web part.** It writes Title, Description,
  Category, Priority, Department and Cc, plus attachments. Ticket ID, Assigned to and
  Status are left to the Power Automate flow; who raised it is SharePoint's Created By.
- Category is built from two questions (`tickets/categories.ts`); the test pins the list
  against the column's 16 choices, so a choice added to the column must be added there.
- Staff choose the priority (it starts at Normal, the column default); IT may change it.
- "My tickets" filters on Created By, so IT staff see only their own there too.
- One web part, two screens: its "Opens on" setting, or `?view=mine` / `?view=raise`.

## Build

SPFx 1.23 needs **Node 22** (not 24). A portable copy lives at
`%USERPROFILE%\tools\node-v22.23.3-win-x64`:

```powershell
$env:Path = "$env:USERPROFILE\tools\node-v22.23.3-win-x64;" + $env:Path
npm install
npm run build   # tests + sharepoint/solution/helpdesk-ticket.sppkg
```

## Deploy

1. Bump `version` in `config/package-solution.json` (both entries).
2. Upload `sharepoint/solution/helpdesk-ticket.sppkg` to the app catalog
   (https://pmwgroupcom.sharepoint.com/sites/appcatalog → Apps for SharePoint), replace, Deploy.
3. Pages already using the web part pick up the new version on their next load.
