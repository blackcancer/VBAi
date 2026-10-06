# Compatibility

[Documentation](README.md)

VBAi targets Windows applications that embed a compatible **64-bit Visual Basic
Editor**. The shared VBIDE layer provides the core editor and assistant;
application adapters provide host document identity, persistence and execution.

## Platform requirements

- Windows, x64 VBE/COM add-in support and .NET Framework 4.8.
- WebView2 Runtime for Monaco; the Windows SQLite runtime for local history.
- The host's required VBA components, references and permitted project access.
- Provider and Git prerequisites only for the features that use them.

The current build has no x86, macOS, Visual Basic 6 IDE or Visual Studio VB.NET
qualification. See [source-build setup](installation.md) for developer prerequisites.

## Shared capabilities and adapter boundaries

| Area | Shared behavior | Host-dependent behavior |
| --- | --- | --- |
| Code | Module/class inspection, editing, import/export, navigation and stale checks | Document-module restrictions, native language services and protection |
| UserForms | Native controls, properties, layout, events and FRM/FRX transport | ActiveX availability, native descriptor behavior and designer initialization |
| Execution/debugging | Guarded project selection, compilation, breakpoint and debugger services | Procedure invocation, runtime inspection and native UI support |
| Persistence | Explicit save requests and verified result reporting | Canonical document identity, save adapters and native prompts |
| Assistant | Context, approvals, tools, review and session history | Successful installed add-in hosting and the requested native operation |

## Observed host qualification

These results are operation-specific and belong to the frozen candidates in
[recorded validation](test-coverage.md). They do not qualify a later rebuild.

| Host | Accepted scope | Principal limits |
| --- | --- | --- |
| Excel x64 | Selected code/debugger/persistence scenarios; complete prepared Q027 UserForm import/recovery/resource-refusal matrix; Q026 options; Q028 embedded Ollama assistant | Other controls, runtime values, DPI/layouts and untested native operations require their own evidence. |
| Word x64 | Adapter save/reopen; canonical document identity; selected Q024 installed Git/chat workflows; Q028 embedded Ollama assistant | Templates, arbitrary SaveAs/cancellation and complete Word UserForm recovery are outside these results. |
| PowerPoint x64 | Selected adapter save/reopen and Q028 embedded Ollama assistant | Broader document, test-explorer and copy-event behavior remains separately qualified. |
| Access x64 | Q012 existing-document adapter/metadata/reference contract; Q028 embedded Ollama assistant | Positive unsupported-Unicode persistence and all designer/runtime behavior are outside acceptance. |
| Publisher x64 | Q012 existing audited publication persistence/metadata/reference contract; Q028 embedded Ollama assistant | NewDocument/first SaveAs and form rendering/events remain unqualified. |
| Classic Outlook x64 | Configured-profile metadata/read-only startup; Q028 synthetic embedded assistant with restored initially empty VBA project | Production OTM files and mail operations are outside qualification. |
| SOLIDWORKS 2019 SP5 and 2025 SP1.1 | Q014 selected native core/editor/designer/local assistant; Q020/Q030 main-desktop native macro creation/publication/reopen | Generic unsafe VBProjects.Open is unavailable; untested controls, signatures and external providers are outside those results. |
| Other VBE hosts | Architectural target where required COM interfaces exist | No acceptance is inferred for Visio, Project or another untested application. |

The Q028 Office matrix uses Office x64 16.0.20430.20092 and the selected
Ollama 0.34.4 CPU/model profile. Other Office builds and provider profiles need
their own campaign. The [release qualification](release-qualification.md) page
records the remaining scope and integration boundaries.

## Document identity and saving

Word document identity is resolved from the PID/COM-identity-matched
`Document.FullName`; raw `VBProject.FileName` can throw or expose temporary
storage. Git bindings refuse stale SaveAs identity. Outlook can expose a fake
absolute project name under different process directories; only a verified
existing `.otm` host path is treated as persisted storage. Otherwise conversation
resolution uses the temporary project identity.

SOLIDWORKS `VBProject.Saved` alone is insufficient persistence evidence. Supported
creation/publication uses fresh paths, verified save bytes and independent native
Edit Macro reopen. Opening a closed form's DesignerWindow can be required before
native property inspection. The unsafe generic `VBProjects.Open` route is refused.

For Access/Publisher Type100 projects, legacy COM HelpFile/HelpContextID writes
are refused before mutation with `SetterStatus=HostLegacyWriteUnsupported`.
Use explicit `read_project_general` and approved `set_project_general` with fresh
identity/revision checks. Representable ANSI text and guarded unsupported-Unicode
refusal have different acceptance. Help-file storage does not qualify opening it.

## Environment-sensitive operations

Native dialogs depend on VBE version, language, DPI and layout. References,
ActiveX availability, protection and trust policy affect individual operations.
Native dark appearance remains experimental. A visible menu or bridge response
alone does not prove persistence, rendering, execution or normal shutdown.

To report another environment, use the host-compatibility issue template. Include
application/VBE/Windows versions, bitness, language, DPI, VBAi candidate identity,
the exact operation and a disposable reproduction. Keep confidential macros out
of reports. See [troubleshooting](troubleshooting.md) for diagnostics.
