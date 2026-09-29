# Updates and the future installer

**The standalone installer is a later deliverable.** VBAi already contains update
coordination and an external updater foundation, but those components are not a
completed installation product. Developer registration is described in
[source-build setup](installation.md).

## Existing update foundation

The update UI handles release discovery, notes, preferences, downloads, ignored
versions and pending-job cancellation. Managed installations can check on startup
and periodically, coordinating across hosts for the same user. Development
checkouts without the installation marker do not apply updates automatically.

The product feed is the VBAi repository, independent of a macro's Git binding.
Public metadata is attempted first; eligible authentication failures can fall back
to the selected Git Credential Manager account without interactive sign-in.
Tokens are not forwarded to the release CDN.

Catalogs and downloads have size and time limits. A compatible installer asset
must provide a SHA-256 digest. The external program rechecks the digest and Windows
Authenticode trust before execution. Source archives are not installer assets.

## Installer contract

A future versioned release supplies `VBAi-Setup-win-x64.msi` or
`VBAi-Setup-win-x64.exe`; the MSI is preferred when both exist. The installer must
have a valid embedded Authenticode signature and deploy the complete add-in, TLB,
resources, dependencies and independent `VBAi.Updater.exe`.

It must register the existing COM identities, preserve user data, handle occupied
files and support rollback. The stable COM assembly identity and the delivery
version are different concerns: MSBuild `ProductVersion` selects the latter.

The installed directory must contain `vbai-installation.json` with product `VBAi`,
architecture `win-x64`, `UpdateProtocol` set to `1` and a persistent, unique
`InstallationId`. A development checkout must not invent this marker to bypass
managed-installation checks.

The updater launches MSI with `/i`, `/quiet`, `/norestart` and logging. An EXE must
implement `/update /quiet /norestart`, wait for the actual installation to finish
and return a reliable exit code. Codes `0`, `3010` and `1641` indicate accepted
success/restart outcomes; after a zero result the installed product version is
checked against the requested target.

## Host lifecycle and recovery

The updater runs independently of the loaded add-in and waits for registered host
processes to exit. PID plus creation time avoids confusing a reused PID with the
original host. It does not terminate applications to force an update.

The future installer must still handle a host starting after that precheck. A
job interrupted after installation begins remains uncertain and is not blindly
re-executed. Pending jobs, host registrations and downloaded payloads live under
`%LOCALAPPDATA%\VBAi\Updates`.

## Prepare a payload, not a release

```powershell
powershell.exe -NoProfile -File tools/Prepare-Release.ps1 -Version 0.1.1-beta.1
```

The version above is an example, not an announced release. The script builds the
Release payload, exports the TLB and prepares packaging metadata. It does not
publish a release or build the future global installer. A supplied `-InstallerPath`
is for an installer already built and signed separately.

The updater also provides `--check-webview2` and `--ensure-webview2` prerequisite
entry points. The latter can download the official Microsoft bootstrapper, verify
its signature and signer, and install it. Having this helper does not establish
that the full installer is ready.

End-to-end signed-release installation, transaction rollback, upgrade, uninstall
and host restarts remain release-qualification work. Do not turn preparation tests
or simulated installer launches into a shipped-installer claim.
