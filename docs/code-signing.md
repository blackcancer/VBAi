# Code signing policy

[Documentation](README.md)

## Current distribution

The maintainer selected an unsigned public 1.0.0 release. Neither Setup nor its
uninstaller has an Authenticode signature. SHA-256 checksums are supplied for
integrity, and release notes disclose this status. The automatic updater's
Windows trust and publisher checks are unchanged: an unsigned package cannot
be applied through it. Do not install a self-signed certificate into a user's
trust stores to bypass these checks.

## Signed build path

`tools/installer/Build-Installer.ps1` uses Inno Setup 6.7 or newer. Signed
compilation requires a provider adapter accepting `-Path`. Inno calls the adapter
for its generated uninstaller and the final setup. The wrapper requires a valid
embedded Authenticode signature and a timestamp before compilation can finish.
Provider credentials belong in a key store or CI secrets, never source files,
compiler arguments or release assets.

The optional `Sign-WithCertificate.ps1` adapter invokes Microsoft SignTool with
SHA-256 file/timestamp digests and an explicit certificate thumbprint. It does
not generate or import certificates. Set `VBAI_SIGNTOOL_PATH`,
`VBAI_SIGNING_CERTIFICATE_THUMBPRINT` and `VBAI_SIGNING_TIMESTAMP_URL` when an
approved certificate is available. Microsoft documents
[SignTool verification](https://learn.microsoft.com/en-us/windows/win32/seccrypto/signtool)
and [timestamping](https://learn.microsoft.com/en-us/windows/win32/seccrypto/time-stamping-authenticode-signatures).

## SignPath Foundation application preparation

No application has been submitted and no signing service has been granted.
The maintainer has selected study of the free open-source route. Before applying:

1. Publish the source and genuine unsigned release with their license/notice scope.
2. Review all bundled dependencies against the Foundation's eligibility rules.
3. Identify actual authors, reviewers and signing approvers; require MFA.
4. Provide the repository, release artifacts, build instructions and privacy policy.
5. Agree with SignPath on source-to-binary verification and the GitHub build chain.
6. Integrate the assigned project, artifact configuration and production signing
   policy after approval, then qualify setup and uninstall signatures separately.

The application entry point is [SignPath Foundation](https://signpath.org/apply.html).
Its [conditions](https://signpath.org/terms.html) require an actively maintained,
documented, released open-source project and define signing roles and provenance
constraints. Acceptance must not be assumed from the chosen MPL/MIT licenses.
The provider signs PE files and scripts; it does not list Inno installers as a
container format for nested signing. Therefore signing the outer EXE alone does
not sign the embedded uninstaller. The Inno signing callback must be connected
to an approved build integration for both files before claiming signed delivery.
See [SignPath artifact formats](https://docs.signpath.io/artifact-configuration/reference)
and [Inno signed uninstallers](https://jrsoftware.org/ishelp/topic_setup_signeduninstaller.htm).

After approval, add the provider's required attribution and the real named roles
to this policy. Until then, VBAi does not claim SignPath sponsorship.
