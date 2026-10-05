# Shared checklist links — design

**Date:** 2026-10-05
**Status:** approved in chat, building

## What it is

IT prepares an asset checklist (IN / OUT / INDIVIDUAL REQUEST) with values
already filled in, then copies a short link and pastes it to the employee
wherever they talk (WhatsApp, Teams, email). The employee opens the link with
no sign-in, fills what is left, signs and submits. From then on the link is a
locked, read-only copy of what was signed, printable to PDF.

"Admin" means anyone signed in to the portal — the portal has no roles.

## What each person sees

**IT, in the portal**

- `/asset-checklist/links` — *Shared checklists*: every link with its employee,
  form type, state (Waiting / Signed / Expired / Cancelled), created and
  expiry dates. Per row: **Copy link**, **QR**, **Open**, **Cancel** (waiting
  only).
- `/asset-checklist/share` — the builder. Pick the form type, fill what is
  known, and beside each field an **Employee can edit** switch. A field left
  blank is always the employee's. Expiry defaults to 14 days. **Create link**
  then shows the link with Copy and a QR code.
- `/asset-checklist` keeps working as today, with a way across to both pages.

**The employee, on `/c/<code>`**

- A separate page (`checklist.html`) carrying only the form kit, the signature
  pad and the logo. No sign-in library, no portal routes, no navigation. Any
  other address still demands a Microsoft sign-in.
- Locked values read as plain text; the rest are inputs. Sign, Submit.
- Afterwards the same link shows the signed record with **Print / Save as
  PDF** (A4 print stylesheet, browser's own PDF).
- Expired, cancelled and unknown codes all read "This link is no longer
  available" — a stranger cannot tell them apart.

## The link

10 characters of base62 from `crypto.getRandomValues` with rejection sampling
(~59 bits). Unguessable within any useful time; together with expiry that is
the protection. The code is the list row's `Title`. The code carries no name,
id or date.

## Storage

New list **Asset Checklist Links**, provisioned by the portal through the
shared `provisionSchema` the first time a link is created:

| Column | Kind | Holds |
|---|---|---|
| Title | text | the code |
| FormMode | choice | In / Out / Individual Request |
| EmployeeName | text | for the IT list only (copy of the preset) |
| Preset | note | JSON of the values IT filled |
| Editable | note | JSON array of fields IT left open |
| ExpiresOn | datetime | |
| LinkStatus | choice | Waiting / Signing / Signed / Cancelled |
| CreatedByName, CreatedByEmail | text | |
| Submitted | note | JSON of the values as signed (no signature) |
| SignatureFile | text | file name in `Signatures` |
| ChecklistId | number | the row in `Asset Checklist Form` |
| SignedOn | datetime | |

The signed checklist itself goes to **`Asset Checklist Form` and
`Signatures` exactly as a portal-signed one does today**, built by the same
`toChecklistItem`. Nothing about existing records or views changes.

## Server

The anonymous browser holds nothing SharePoint accepts, so two Vercel
functions do the reading and writing under the portal's own app identity:

- `GET /api/c/:code` — `{ state: 'open', formMode, values, editable, expiresOn }`,
  or `{ state: 'signed', values, signature, signedOn }`, or `404 { state: 'gone' }`.
- `POST /api/c/:code` with `{ values }`:
  1. read the link; refuse unless open;
  2. **claim** it — `LinkStatus = Signing` with `If-Match` on the item's eTag;
     a conflict means somebody else is submitting and is refused;
  3. clean the submission (`cleanSubmission`: known options only, lengths and
     quantities bounded, signature must be a PNG data URL under 1 MB);
  4. merge (`mergeSubmission`): locked fields come from the preset whatever the
     browser sent, `formMode` always from the link;
  5. validate with the existing `validateChecklist`;
  6. upload the signature, write the checklist row, then mark the link Signed
     with `Submitted`, `SignatureFile`, `ChecklistId`, `SignedOn`;
  7. on any failure after the claim, put the link back to Waiting. A Signing
     state older than 5 minutes is read as Waiting, so a crashed function
     cannot strand a link.

Talks to SharePoint through **Microsoft Graph** (`Sites.Selected`, client
secret): SharePoint's own REST API rejects app-only tokens obtained with a
secret and would need a certificate. Field names are the same internal names
the portal writes.

Every answer carries `Cache-Control: no-store`, `X-Robots-Tag: noindex`,
`Referrer-Policy: no-referrer`; `/c/*` gets the same headers in `vercel.json`.

Creating and cancelling links is done by the portal as the signed-in user,
over SharePoint REST, like every other write in the portal.

## Code layout

- `src/features/forms/links/` — pure: `linkCode.js`, `linkRules.js`
  (`employeeMayEdit`, `editableFields`, `linkState`, `mergeSubmission`,
  `cleanSubmission`), `linkSchema.js`. Tested.
- `src/features/forms/sharepoint/checklistLinks.js` — portal-side create,
  list, cancel (REST).
- `server/` — `graph.js` (token + fetch), `checklistLinkApi.js` (the two
  handlers over an injected Graph client, tested with a fake).
- `api/c/[code].js` — the Vercel entry, wiring env to the handlers.
- `src/components/checklist/ChecklistFields.jsx` — the details/items/serials/
  remarks block, extracted from `AssetChecklistPage` so the portal page, the
  builder and the public page draw the same form. Takes `locked` and an
  optional per-field adornment (the builder's switch).
- `src/components/checklist/ChecklistRecord.jsx` — the read-only signed copy.
- `checklist.html` + `src/public/checklist.jsx` — the public entry.
- Vite: second build input; a dev middleware serves `/c/*` from
  `checklist.html` and `/api/c/*` from the same handlers.

## Setup the admin does once

1. Azure → App registrations → new app *PMW IT Portal — server*.
2. API permissions → Microsoft Graph → Application → `Sites.Selected`; grant
   admin consent.
3. Give it `write` on the IThelpdesk site (Graph
   `POST /sites/{site-id}/permissions`, or PnP
   `Grant-PnPAzureADAppSitePermission`).
4. Certificates & secrets → new client secret.
5. Vercel env: `CHECKLIST_APP_CLIENT_ID`, `CHECKLIST_APP_CLIENT_SECRET`,
   `CHECKLIST_TENANT_ID` (falls back to `VITE_TENANT_ID`), and the site URL from
   `VITE_SHAREPOINT_SITE_URL`.

Without these the public page says "This form cannot be opened right now" and
the server logs which setting is missing.

## Out of scope

Emailing the link, reminders, editing a link after creation (cancel and make
a new one), a server-rendered PDF, per-IP rate limiting (serverless instances
share no memory; the code's entropy is the protection).

## Testing

Unit: code generation alphabet/length; `employeeMayEdit` (blank preset is
always editable); `mergeSubmission` cannot override a locked field or the
mode; `cleanSubmission` drops unknown options and oversize input; `linkState`
for expired / cancelled / signed / stale Signing; the handlers against a fake
Graph — refuse closed links, refuse on claim conflict, roll back on a failed
write, write the same row `toChecklistItem` produces. Then the builder and the
public page end to end in the dev server against a fake Graph.
