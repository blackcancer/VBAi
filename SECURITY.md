# Security policy

## Report vulnerabilities privately

Email **init-sys-rev@hotmail.com** with the subject `VBAi security report`.
This is the repository's existing maintainer contact. Do not use a public issue
or the in-app GitHub problem-report action for an undisclosed vulnerability.

Include the affected commit or build, application and VBE versions, architecture,
required permissions, a minimal reproduction, observed impact and any suggested
mitigation. Use disposable projects and redact credentials, private paths and
business data. Start with a description rather than a large confidential attachment.

Please coordinate disclosure while the report is assessed. The project does not
currently promise a response deadline, a bug bounty or paid incident response.

## Maintenance scope

VBAi is under active development. Report against the latest relevant development
revision, and mention whether an older build is also affected. A versioned security
support window has not been established; historic builds are not implicitly
maintained indefinitely.

## Security boundaries

VBAi runs inside a desktop host process. Its conversation permissions, tool
approval policies and revision checks are application controls, not an operating
system sandbox. Executing VBA can affect documents, files, services and other
resources accessible to the host.

The chat limits project access and separately controls shared VBE context. The
local diagnostic bridge is a different surface with Windows-account access
controls; it is not a public remote API and is not governed by chat approvals.
See [privacy and safety](docs/privacy.md) and [architecture](docs/architecture.md).

Credentials and editor drafts use Windows DPAPI. This does not imply that every
local file is encrypted, nor does it protect data after it has been sent to an
external provider. Treat project content and tool results as untrusted input.

## Safer use

Work on backups before allowing edits or execution. Review the provider and
context permissions, and follow your organization's data-sharing policy. Inspect
changes before running a macro. Do not disable host security controls globally
or install an untrusted component to bypass an unsupported operation.

The standalone installer is a future deliverable. Release-signature requirements
in [the update contract](docs/updates.md) describe the intended distribution path,
not a claim that a signed installer is already available.
