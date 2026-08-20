# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-04
**Updated:** 2026-08-20
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
│   ├── hooks/
│   │   ├── useRequests.js    # the one SharePoint read + row helpers + token
│   │   └── useSession.js     # session context + phases (no component here)
│   ├── pages/                # Homepage, LoginPage, DashboardPage, ListPage,
│   │                         # FormPage, AssetChecklistPage
│   ├── styles/
│   │   ├── shell.css         # tokens, brand surface, shell, UI, dashboard
│   │   └── auth.css          # sign-in layout + the idle animation
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
