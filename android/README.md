# PMW IT — Android app

The portal as an installable Android app. It is a **Trusted Web Activity**:
Chrome on the phone shows the live site (`https://pmw-it.vercel.app`) full
screen, with no address bar. Sign-in, the camera at the till and everything
else behave exactly as they do in Chrome.

## Updates

None of the portal is inside the APK, so **a push to main reaches every phone
without a new APK**:

- Opening the app always loads the newest deploy.
- **Check for updates** sits in the menu, above Sign out, in the app only (not in
  a browser tab). A newer build loads at once when pressed.
- On launch it checks straight away and, if behind, reloads into the new build
  before anything is typed (once per build, so a slow CDN cannot loop it).
- An app left open checks every five minutes and whenever it comes back to the
  foreground (`src/components/UpdateWatcher.jsx`). When a newer deploy exists it
  shows "A new version of the portal is ready", and the next move to another
  screen reloads into it. It never reloads under somebody mid-form on its own;
  "Update now" does it at once.

Rebuild the APK only to change its name, icon, colours or address.

## Installing

`app/build/outputs/apk/release/app-release.apk` — send it to the phone (or push
it through Intune) and allow "install unknown apps" for whatever opens it.
The phone needs Chrome, which company Android phones already have.

The full-screen look depends on `public/.well-known/assetlinks.json` being live
on the site. Until that deploy is out, the app still works but shows a thin
address bar at the top.

## Building

Needs JDK 17 and the Android SDK (platform 34). From this folder:

    gradlew assembleRelease

Another address: `gradlew assembleRelease -PtwaHost=it.example.com` (and put
`assetlinks.json` on that site too).

Signing uses `signing/keystore.properties` and `signing/pmw-it-release.jks`,
which are NOT in git. See `signing/README.md` and back them up.

New icons from the logo: `pwsh ../scripts/make-app-icons.ps1`.

### "Plugin com.android.application was not found"

This PC's Norton antivirus re-signs HTTPS traffic, and Java only trusts it if
Norton's CURRENT root certificate is in the trust store Java is told to use
(`JAVA_TOOL_OPTIONS`). The copy in `C:\Android\certs\gradle-truststore.jks`
predates the one Norton uses now. Export the current one from Windows
(certmgr → Trusted Root → "Norton Web/Mail Shield Root") and import it:

    keytool -importcert -keystore C:\Android\certs\gradle-truststore.jks -storepass changeit -alias norton-current -file norton.cer
