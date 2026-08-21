# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-04
**Updated:** 2026-08-21
**Project:** PMW IT Service Portal (formerly "IT Onboarding Portal")

## OVERVIEW
React 19 + Vite 8 SPA with Azure AD MSAL authentication and SurveyJS forms. Deployed on Vercel.

The UI is the SI CMMS shell — a branded nav column, a sticky bar, a dashboard of
stat cards over a canvas — minus everything maintenance-specific (work orders,
machines, priorities, SLA, roles). The sign-in screen is the PMW OSHE portal's
split poster/card layout, with an idle animation of its own.

## STRUCTURE
```
pmw-it/
├── src/
│   ├── components/
│   │   ├── AppShell.jsx      # nav column + sticky bar + main; the one auth gate
│   │   ├── IdleAnimation.jsx # the sign-in screen's chip/packet animation
│   │   ├── Logo.jsx          # PMW mark (src/assets/logo-*.png)
│   │   ├── SessionDialog.jsx # what a timed-out session looks like being fixed
│   │   ├── SignInTransition.jsx # post-sign-in veil, fades into the dashboard
│   │   ├── SignatureDialog.jsx
│   │   └── ui/               # Icons, Button, Surfaces, StatCard, Badges
│   ├── features/
│   │   ├── datastudio/       # Excel import + profiling (in progress)
│   │   └── devices/          # parse/ derive/ sharepoint/ stats/ ui/
│   ├── hooks/
│   │   ├── useRequests.js    # the one SharePoint read + row helpers + token
│   │   └── useSession.js     # session context + phases (no component here)
│   ├── pages/                # Homepage, LoginPage, DashboardPage, ListPage,
│   │                         # FormPage, AssetChecklistPage, DevicesPage
│   ├── styles/
│   │   ├── shell.css         # tokens, brand surface, shell, UI, dashboard
│   │   ├── auth.css          # sign-in layout + the idle animation
│   │   └── devices.css       # the device list section
│   ├── context/              # ThemeContext (dark/light), SessionContext (auto
│   │                         # re-sign-in + its dialog and entrance animation)
│   ├── services/             # sharePointService.js
│   ├── utils/                # timeout.js, initials.js, authErrors.js,
│   │                         # sessionKeys.js
│   ├── App.jsx               # Router setup
│   ├── main.jsx              # MSAL bootstrap + providers + stylesheet order
│   └── authConfig.js         # Azure AD + SharePoint scopes
├── public/                   # Static assets
├── vite.config.js
├── eslint.config.js
└── package.json
```

## ROUTES
| Path | Screen |
|------|--------|
| `/` | Redirects to `/dashboard` or `/login` |
| `/login` | Sign-in poster + card |
| `/dashboard` | Stat cards, charts, latest requests |
| `/requests` | The records table; filters live in the query string |
| `/list` | Legacy alias, redirects to `/requests` (keeps the query) |
| `/it-boarding-form` | SurveyJS request form (`?edit=<id>` opens a record) |
| `/asset-checklist` | Handover checklist (IN / OUT / individual) |
| `/devices` | Device list: fleet dashboard, register and scan-report import (`?view=`) |

## WHERE TO LOOK
| Task | Location |
|------|----------|
| Auth logic | `src/main.jsx` (MSAL init), `src/authConfig.js` |
| Auth gate for a page | `src/components/AppShell.jsx` — pages do not gate themselves |
| Timed-out session / auto sign-in | `src/context/SessionContext.jsx` |
| What counts as a dead session | `isInteractionRequired` in `src/utils/authErrors.js` |
| Any SharePoint token | `useSharePointToken()` in `src/hooks/useRequests.js` |
| Routes | `src/App.jsx` |
| Nav items | `NAV_ITEMS` in `src/components/AppShell.jsx` |
| Design tokens / layout | `src/styles/shell.css` |
| Sign-in screen | `src/pages/LoginPage.jsx`, `src/styles/auth.css` |
| SharePoint reads | `src/hooks/useRequests.js` |
| Device report parsing | `src/features/devices/parse/` |
| Device derived fields and risk | `src/features/devices/derive/` |
| Device SharePoint schema | `src/features/devices/sharepoint/deviceSchema.js` |
| Device SharePoint list views | `src/features/devices/sharepoint/deviceViews.js` |
| Editing or removing one device row | `src/features/devices/sharepoint/updateDevice.js` |
| Device fleet statistics | `src/features/devices/stats/deviceStats.js` |
| Bar and column charts | `src/components/ui/Charts.jsx` (shared by both dashboards) |
| SharePoint writes | `src/services/sharePointService.js` |
| Theme | `src/context/ThemeContext.jsx`; toggle lives in the shell's bar |

## CONVENTIONS

**Page composition**: a screen renders `<AppShell title subtitle actions>` and
its own body. The bar, the nav, the theme toggle, sign-out and the sign-in gate
all belong to the shell — do not re-add per-page copies of them.

**Stylesheet order** (`src/main.jsx`): `index.css` → `App.css` → `styles/shell.css`
→ `styles/auth.css`. shell.css re-points the older `--bg` / `--surface` /
`--border` / `--text-*` tokens at the new palette, so it must load last.
`--accent` is deliberately left alone: it fills `.ms-button`, whose text colour
is `--bg`.

**Dashboard ↔ records**: every dashboard figure links into `/requests` with a
query string (`?type=`, `?entity=`, `?department=`, `?range=`, `?equipment=`).
Both screens read the same `useRequests()` fetch, so a card and the list it opens
cannot disagree.

**Navigation**: use `window.location.replace()` instead of React Router
`navigate()` *inside `useEffect`*. WHY: navigate causes a state update →
re-render → effect runs again → infinite loop. In event handlers `navigate()` is
correct and is what the shell and pages use.

**MSAL redirect handling**: always await `handleRedirectPromise()` before
rendering. Silent `no_token_request_cache_error` is normal on fresh load.

**Session guard** (`src/context/SessionContext.jsx`): a timed-out session signs
itself back in. Two rules keep it from disturbing anyone who is still signed in:

1. Nothing is on a clock — no idle timer, no expiry watcher. A recovery starts
   only on proof (an `InteractionRequiredAuthError`, or MSAL's silent renewal
   coming back `timed_out`), or when the account has vanished from the cache on
   a page that needs one. Network failures and SharePoint errors stay errors.
2. A recovery tries `ssoSilent` first. A live Azure AD session comes back in
   about a second and nobody leaves the page; only a second refusal escalates to
   `loginRedirect`, guarded by a one-shot `sessionStorage` flag so a sign-in that
   keeps failing cannot bounce the browser in a loop.

`/login` is exempt, and sign-out clears the stored login hint *before* the
redirect — leaving it would let the guard read a deliberate sign-out as a
timeout and undo it. Sign out through `useSession().signOut()` for that reason.

**`src/features/<name>/`** is where a section with more than a handful of modules
lives — `datastudio/` and `devices/` both follow it. Layering inside a feature:
`parse/` knows nothing about the domain, `derive/` knows nothing about SharePoint,
`sharepoint/` imports no React. Each layer is testable without the one above it.

**Device report parsing keys off a known-label whitelist** (`parse/labels.js`). A
generic `^Word:` split reads `Total Slots: 2 | Used Slots: 2` and
`Y: | \\server\PMW\IT` as field names and moves those values out of the blocks they
belong to. An unknown label owns the lines beneath it, so a field the scan script
adds later surfaces in review rather than contaminating its predecessor.

**A hand-edited device field outranks the scan file.** The register lets the
three DERIVED fields be retyped (owner, department, device type) and records
which ones in `ManualFields`. `applyManualOverrides` in `syncDevices.js` then
holds those back on re-import — from the diff AND from the body, or updating
anything else would overwrite them as a side effect. Clearing a field is how
somebody hands it back to the scan file; without that it would stay frozen
against every future import for good.

**`Total RAM` in a scan report is usable RAM, not installed RAM.** Windows subtracts
the integrated GPU's reserved share, so a 16 GB laptop reports 15 GB and an 8 GB one
reports 7 GB. Sum `RAM Slot Info` for the real figure; ranking on the reported one
puts a 16 GB machine below an 8 GB machine.

**SharePoint column creation**, verified against the tenant on 2026-08-21 while
provisioning the device lists. All three of these fail silently or confusingly
if you get them wrong:

1. **Create each field as its concrete type**, not the base `SP.Field`.
   `SP.Field` does not declare `Choices`, so a choice column sent that way
   fails with *"The property 'Choices' does not exist on type 'SP.Field'"*.
   Use `SP.FieldChoice`, `SP.FieldNumber`, `SP.FieldDateTime`,
   `SP.FieldMultiLineText`.
2. **The internal name comes from the `Title` a field is CREATED with.**
   `StaticName` in the creation body does not control it. Create
   `Title: 'Device Type'` and the field is addressable only as
   `Device_x0020_Type`; every item write of `DeviceType` then fails with
   *"The property 'DeviceType' does not exist on type 'SP.Data...ListItem'"*.
   Create under the internal name, then MERGE the display `Title` on
   afterwards. This is what produced the hand-encoded `Calling_x0020_Name`
   in `sharePointService.js`.
3. **Read `InternalName`, never `StaticName`**, when checking which columns
   already exist. The two can disagree, and a column where they disagree is
   precisely the broken one.
4. **REST-created columns join no view.** A freshly provisioned list shows
   nothing but its Title until view fields are set explicitly, which is what
   `deviceViews.js` is for.
5. **`ViewQuery` is only honoured in the creation body.** A default view is
   never created, so a filter or sort declared on one has to be MERGEd on
   afterwards or it is silently dropped. Address the built-in view through
   `/defaultView`, not `getByTitle('All Items')` — that title is English-only.

`ensureAssetColumns` in `src/services/sharePointService.js` still has bug 1:
it sends `Choices` with `__metadata: SP.Field`. Its lists predate the bug, so
nothing is broken today, but the same code on a fresh site would fail.

**SharePoint scopes**: use ROOT domain only, never site paths.
- ✅ `https://pmwgroupcom.sharepoint.com/AllSites.Write`
- ❌ `https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/AllSites.Write`

**Icons**: `src/components/ui/Icons.jsx`, transcribed on the 24px stroke grid.
No icon package is installed — add a glyph there rather than a dependency.

## ANTI-PATTERNS (THIS PROJECT)
- Don't use `navigate()` in useEffect — causes infinite loops
- Don't add SharePoint scopes to loginRequest — separate request required
- Don't call `acquireTokenSilent` / `acquireTokenPopup` from a page. Use
  `useSharePointToken()`, which routes through the session guard. The popup
  fallback this replaced is where the "it just stopped loading" reports came
  from: a popup opened from an expired timer is not a user gesture, so browsers
  block it and the page waits forever on a window nobody was shown.
- Don't widen `isInteractionRequired` to catch more errors. Signing someone back
  in who never lost their session is worse than the error they were going to
  see; ambiguous codes belong on the "not a timeout" side.
- Don't use `useNavigate` for auth redirects — use window.location
- Don't add `assetsInclude: ['**/*.html']` to vite.config.js. It matches
  index.html itself, so Vite stops treating it as the HTML entry and emits it as
  an asset — `npm run build` then produces a dist/index.html containing
  `export default "/assets/index-….html"` and no bundle at all.
- Don't create a SharePoint DateTime column with `DisplayFormat: 0` when the time
  matters — that is DateOnly and silently discards it. Device columns use `1`,
  confirmed round-tripping a real instant back out of the list.
- Don't give `.bar-fill` anything but `display: block` in `shell.css`. It is a
  `<span>` inside `.bar-track`, which is a plain block, so it is not blockified
  the way a flex or grid child would be — left inline it ignores width and
  height and every dashboard bar paints as an empty track.
- Don't create a SharePoint Note column without `RichText: false`; a rich-text Note
  wraps stored values in `<div>` markup and will not round-trip.
- Don't send `Choices` on a base `SP.Field`. A property exists only on the type
  that declares it, and the tenant answers "The property 'Choices' does not exist
  on type 'SP.Field'". Choice columns go out as `SP.FieldChoice`.
- Don't create a SharePoint column under its display name. SharePoint derives the
  internal name from the `Title` a field is created with — `StaticName` in the
  creation body does not control it — so "Device Type" becomes
  `Device_x0020_Type` and every item write then fails with "The property
  'DeviceType' does not exist". Create the column under its internal name and
  MERGE the display name onto it afterwards, as `provisionLists.js` does. The
  hand-encoded `Calling_x0020_Name` in `sharePointService.js` is the same trap,
  paid the other way.
- Don't add `hour12` beside `hourCycle` in `malaysiaTime.js` — per the Intl spec an
  explicit `hour12` nullifies `hourCycle` entirely. The 24-hour path pins `h23`, the
  AM/PM path pins `h12`, and neither passes `hour12`.
- Don't export a helper next to a component from the same file — it drops the
  file out of Fast Refresh (and eslint fails the build). `initialsOf` lives in
  `src/utils/initials.js` for exactly this reason.

## COMMANDS
```bash
npm run dev      # Start dev server on port 5173
npm run build    # Build for production (outputs to dist/)
npm run lint     # Run ESLint
npm run preview  # Preview production build
```

## NOTES
- Vite port 5173 is for local dev; Vercel ignores this
- MSAL handles Azure AD login flow + token caching
- SurveyJS drives `/it-boarding-form` and `/asset-checklist`
- `npm run lint` still reports pre-existing errors in FormPage,
  AssetChecklistPage, SignatureDialog and ThemeContext (unused imports, SurveyJS
  model mutation inside hooks). They predate the shell work and are untouched.
