# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-04
**Project:** PMW IT Onboarding Portal

## OVERVIEW
React 19 + Vite 8 SPA with Azure AD MSAL authentication and SurveyJS forms. Deployed on Vercel.

## STRUCTURE
```
pmw-it/
├── src/
│   ├── pages/          # Route components (Homepage, FormPage, ListPage, LoginPage)
│   ├── context/        # ThemeContext (dark/light mode)
│   ├── services/       # sharePointService.js
│   ├── App.jsx         # Router setup
│   ├── main.jsx        # MSAL bootstrap + providers
│   └── authConfig.js   # Azure AD + SharePoint scopes
├── public/             # Static assets
├── vite.config.js      # Vite config
├── eslint.config.js    # ESLint flat config
└── package.json
```

## WHERE TO LOOK
| Task | Location |
|------|----------|
| Auth logic | `src/main.jsx` (MSAL init), `src/authConfig.js` |
| Routes | `src/App.jsx` |
| Pages | `src/pages/*.jsx` |
| SharePoint integration | `src/services/sharePointService.js` |
| Theme | `src/context/ThemeContext.jsx` |

## CONVENTIONS

**Navigation**: Use `window.location.replace()` instead of React Router `navigate()` in `useEffect`. WHY: navigate causes state update → re-render → effect runs again → infinite loop.

**MSAL redirect handling**: Always await `handleRedirectPromise()` before rendering. Silent `no_token_request_cache_error` is normal on fresh load.

**SharePoint scopes**: Use ROOT domain only, never site paths.
- ✅ `https://pmwgroupcom.sharepoint.com/AllSites.Write`
- ❌ `https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/AllSites.Write`

## ANTI-PATTERNS (THIS PROJECT)
- Don't use `navigate()` in useEffect - causes infinite loops
- Don't add SharePoint scopes to loginRequest - separate request required
- Don't use `useNavigate` for auth redirects - use window.location

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
- SurveyJS used for `it-boarding-form` page (forms library)
- QR code library available for future use