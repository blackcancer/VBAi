# Documentation

VBAi is a development workspace for applications that host the Visual Basic Editor.
Start with the task you want to complete; the host compatibility layer is documented
separately from the shared VBE features.

## Use VBAi

| Guide | What it answers |
| --- | --- |
| [Getting started](getting-started.md) | How do I complete a first, controlled intervention? |
| [Compatibility](compatibility.md) | What does the shared VBE layer cover, and what depends on the application? |
| [Source-build setup](installation.md) | How do developers and early testers build and register the current preview? |
| [Conversations](chat-ui.md) | How do context, modes, approvals, queues, recovery and session history work? |
| [Providers](providers.md) | How do I configure authentication, endpoints and models? |
| [Modern editor](modern-editor.md) | How do editing, synchronization, saving and native debugging interact? |
| [Git and GitHub](github-integration.md) | How do source versioning, checkpoints and imports work? |
| [UserForms](reference/designer.md) | How do native forms, controls and resource files behave? |
| [Privacy and safety](privacy.md) | What is shared, what stays local, and what permissions mean? |
| [Troubleshooting](troubleshooting.md) | How do I recover from setup, editing, provider or UI problems? |

## Develop and maintain

| Guide | Scope |
| --- | --- |
| [Architecture](architecture.md) | Components, host boundaries, threading, local bridge and persistence. |
| [Development](development.md) | Build conventions, UI design, localization and documentation maintenance. |
| [Testing](../tests/README.md) | Local checks and explicit real-host test opt-ins. |
| [Recorded validation](test-coverage.md) | Dated evidence, tested revisions and measurement boundaries. |
| [Version 1.0.0 qualification](release-qualification.md) | Release gates, tracked defects and remaining native acceptance work. |
| [Tool reference](reference/vbe-tools.md) | Discovery, invocation and authoritative schema locations. |
| [Updates and release contract](updates.md) | Existing updater foundation and requirements for the future installer. |
| [Roadmap](roadmap.md) | Planned work without delivery promises. |

Repository policies: [contributing](../CONTRIBUTING.md),
[security](../SECURITY.md), [support](../SUPPORT.md),
[conduct](../CODE_OF_CONDUCT.md), [third-party notices](../THIRD_PARTY_NOTICES.md)
and [changelog](../CHANGELOG.md).

## Read status accurately

**Implemented** describes code that exists. **Tested** describes an identified
scenario on a particular build and environment. **Not run**, **blocked** and
**not installed** are not successful tests. A coverage percentage measures only
the instrumented scope.

Guides describe the maintained behavior. Detailed investigation logs, superseded
plans and old test counters remain in Git history rather than a parallel archive
of competing documentation. Non-Markdown reference data is retained where useful.
