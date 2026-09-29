# VBE compatibility and host adapters

## Product scope

VBAi targets the **Visual Basic Editor embedded in Windows applications**. It is
not an Excel-only or SOLIDWORKS-only add-in. These applications are examples of
VBE hosts and test environments, not the product's architectural boundary.

The current build targets **x64** and **.NET Framework 4.8**. The application must
expose a compatible VBE/COM add-in environment. Monaco uses WebView2; chat storage
uses the Windows SQLite runtime. This build does not claim x86, macOS, the Visual
Basic 6 IDE or Visual Studio's VB.NET editor as supported targets.

## Three different compatibility questions

| Layer | Responsibility | What must be established |
| --- | --- | --- |
| Shared VBE/VBIDE integration | Projects, modules, references, forms, code panes, native editor operations and the assistant workspace. | The host exposes the required interfaces and accepts the add-in. |
| Host-specific compatibility | Document identity, persistence, execution and application-specific behavior not provided by VBIDE alone. | An adapter exists for the requested operation and handles this host's behavior. |
| Qualification | Observed results in a specific application/runtime/environment. | A reproducible test passed on the identified build. |

A missing save adapter does not negate working shared editor capabilities. Equally,
a host that loads the add-in is not thereby qualified for every mutation, debugger
operation, language feature or UI surface.

## Recorded host evidence

This table summarizes evidence available at documentation baseline
`99b5f25` on **2026-09-29**. These are recorded project results, not new tests run
as part of the documentation rewrite. Full provenance is in
[recorded validation](test-coverage.md).

| Host/environment | Observed scope | Important boundaries |
| --- | --- | --- |
| Excel x64 | Multiple recorded load, bridge, editing, persistence, UserForm and Monaco scenarios; separate native compilation/debugging checks. | Individual runs cover different features and revisions, not every combination. |
| Word, Office 16 x64 | Disposable `.docm`: shared VBE inspection, module/class/form changes, compilation, native-helper save and reopen. | VBAi's save route was refused because programmatic project access was not trusted on the test machine. `Normal` was not modified. |
| PowerPoint, Office 16 x64 | Disposable `.pptm`: the same shared VBE scenario set and native-helper persistence. | The VBAi save adapter attempted to read unavailable `Application.HWND`; that operation remains a known defect in this evidence. |
| Access, Office 16 x64 | Disposable `.accdb`: shared VBE scenarios, including an MSForms UserForm, and native-helper persistence. | VBAi document-save adapter absent; refusal verified. |
| Publisher, Office 16 x64 | Disposable `.pub`: shared VBE scenarios and native-helper persistence. | VBAi document-save adapter absent; refusal verified. |
| Classic Outlook | A read-only qualification path is prepared. | Blocked by first-run setup and the absence of a configured classic Outlook profile. No mail or user VBA project was modified. |
| Visio | No result in the current evidence set. | Not installed on the qualification machine. |
| SOLIDWORKS 2019 SP5 | Historical load, compile/run, breakpoint, stepping and resume checks on a disposable macro. | Later Monaco and standalone `.swp` save paths are not covered by that historical result; reading local variable values remains unqualified. |
| Other applications with a compatible VBE | Within the intended product scope. | Require load testing and operation-specific qualification; do not infer a result from another host. |

**Saving through a test helper's application API does not validate VBAi's
`save_host_document` tool.** Native VBE Save, an application adapter and the host's
own save API are distinct paths.

The Office batch did not qualify macro execution, stepping, signatures, all
controls, the entire theme/Monaco UI or LLM providers in every application.

## Cross-host constraints

Native dialogs and accessibility depend on VBE version, UI language, DPI and window
layout. COM references, ActiveX availability, protected projects and trust settings
also affect individual operations. Third-party controls require their own evidence.
The native dark theme remains experimental and uses additional OS-dependent
behavior; it is not part of a universal host-compatibility guarantee.

## Report another environment

Use the host-compatibility issue template. Record application/version, VBE version,
Windows build, process architecture, UI language, DPI, VBAi commit and the exact
operation. Distinguish **pass**, **failed**, **blocked by a prerequisite** and
**not run**. A minimal disposable file and steps are more useful than a blanket
“works” or “does not work” report. Never submit confidential production documents.
