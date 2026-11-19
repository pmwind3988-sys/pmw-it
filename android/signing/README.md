# Signing key — keep a copy somewhere safe

`pmw-it-release.jks` and `keystore.properties` (its passwords) sign the APK.
They are deliberately not in git.

- **Back both files up** (password manager / IT share). Android only installs an
  update over the app if it is signed with the SAME key. Lose it and every phone
  has to uninstall the app and install a new one.
- The key's SHA-256 fingerprint is in `public/.well-known/assetlinks.json`.
  That file is what lets the app open the portal full screen, with no address
  bar. A new key means a new fingerprint there.

Print the fingerprint again:

    keytool -list -v -keystore pmw-it-release.jks -alias pmw-it
