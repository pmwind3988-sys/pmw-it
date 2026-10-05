# Shared Checklist Links Implementation Plan

> Executed inline by the author of the spec in the same session, so steps name
> files, interfaces and test cases rather than repeating the code.

**Goal:** IT copies a short link to a pre-filled asset checklist; the employee
fills, signs and submits it without signing in; the link then shows a locked,
printable copy.

**Architecture:** Pure rules in `src/features/forms/links/`, shared by the
portal, the public page and two Vercel functions. The functions reach
SharePoint through Graph with an app-only token; the portal creates and
cancels links over SharePoint REST as the signed-in user.

**Tech Stack:** React 19, Vite 8 (second HTML entry), Vercel Node functions,
Microsoft Graph, Vitest.

**Spec:** `docs/superpowers/specs/2026-10-05-shared-checklist-links-design.md`

## Global Constraints

- No new npm dependency (`qrcode` is already installed).
- Stored `FormMode` values stay `In` / `Out` / `Individual Request`.
- Signed rows go to `Asset Checklist Form` via `toChecklistItem`, unchanged.
- Public page imports nothing from MSAL, `AppShell`, routes or `useRequests`.
- Choice columns are only ever added to.

---

### Task 1: Pure link rules

**Files:** create `src/features/forms/links/linkCode.js`, `linkRules.js`,
`linkSchema.js`, `links.test.js`.

**Produces:**
- `newLinkCode(random = crypto.getRandomValues) → string` (10 chars base62),
  `isLinkCode(s) → boolean`.
- `LOCKABLE_FIELDS`, `isBlankValue(field, value)`,
  `employeeMayEdit(link, field)`, `editableFields(link) → string[]`.
- `linkState(link, now) → 'open'|'busy'|'signed'|'expired'|'cancelled'`.
- `cleanSubmission(values) → values` (known options, bounded lengths,
  quantities 1–999, PNG data URL ≤ 1 MB else null).
- `mergeSubmission(link, submitted) → values`.
- `LINKS_LIST_NAME`, `LINK_COLUMNS`, `LINK_VIEWS`, `LINK_STATUS`,
  `toLinkItem(...)`, `fromLinkItem(fields) → link`.

**Tests:** code alphabet/length/no modulo bias path; blank preset editable;
locked field and mode cannot be overridden; unknown option dropped; oversize
signature rejected; expired/cancelled/signed/stale-Signing states.

### Task 2: Server handlers over a fake Graph

**Files:** create `server/graph.js`, `server/checklistLinkApi.js`,
`server/checklistLinkApi.test.js`, `api/c/[code].js`; vitest include `server/`.

**Produces:** `createLinkApi({ graph, now })` →
`{ get(code), submit(code, body) }` returning `{ status, body }`;
`createGraph({ tenantId, clientId, clientSecret, siteUrl })`.

**Tests:** gone for unknown/expired/cancelled; open returns preset + editable;
submit refuses closed, refuses claim conflict (412), rolls back on failed row
write, writes `toChecklistItem` output, marks Signed with ChecklistId; signed
GET returns submitted values + signature.

### Task 3: Shared form block

**Files:** create `src/components/checklist/ChecklistFields.jsx`,
`ChecklistRecord.jsx`; modify `src/components/form/Field.jsx` (adornment),
`src/pages/AssetChecklistPage.jsx` (use the block), `src/styles/forms.css`.

### Task 4: Portal — builder and list

**Files:** create `src/features/forms/sharepoint/checklistLinks.js`,
`src/pages/ChecklistSharePage.jsx`, `src/pages/ChecklistLinksPage.jsx`;
modify `src/App.jsx`, `AssetChecklistPage.jsx` (links across).

### Task 5: Public page, build and deploy wiring

**Files:** create `checklist.html`, `src/public/checklist.jsx`,
`src/public/PublicChecklist.jsx`, `src/styles/public.css`; modify
`vite.config.js` (input + dev middleware), `vercel.json`.

### Task 6: Verify and document

Run tests, lint, build; drive builder + public page in the dev server against
a fake Graph; update `AGENTS.md` (routes, where-to-look, conventions) and add
`docs/checklist-links-setup.md` for the admin.
