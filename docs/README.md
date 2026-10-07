# Documentation

Use these guides to set up VBAi, complete a workflow or contribute to the code.
The product targets the shared VBE environment; host qualification is documented
per operation and tested candidate.

The [public website](https://vbai.app/) presents installation,
AI connections and a real illustrated workflow in English and French. Its
[maintenance guide](../site/README.md) explains sources and publication.

## Start here

| Your task | Guide |
| --- | --- |
| Install VBAi or build from source | [Installation](installation.md) |
| Complete a first intervention | [Getting started](getting-started.md) |
| Read the French offline manual | [User manual](help/README.md) |
| Check your application | [Compatibility](compatibility.md) |
| Connect an AI provider | [Providers](providers.md) |
| Solve a setup/runtime problem | [Troubleshooting](troubleshooting.md) |

## Work with VBAi

| Guide | Covers |
| --- | --- |
| [Conversations](chat-ui.md) | Context, modes, approvals, queues, activities, recovery and history |
| [Modern editor](modern-editor.md) | Completion, synchronization, drafts, native compilation and debugging |
| [UserForms](reference/designer.md) | Controls, typed properties, layout, events and FRM/FRX resources |
| [Git and GitHub](github-integration.md) | Source repositories, checkpoints, imports and remote workflows |
| [VBA test explorer](vba-testing.md) | Annotations, serial runs, results and procedure-entry coverage |
| [Privacy and safety](privacy.md) | Project boundaries, local storage, provider transmission and execution |

## Develop and maintain

| Guide | Covers |
| --- | --- |
| [Architecture](architecture.md) | Components, COM/threading boundaries and dispatch lifecycle |
| [Development](development.md) | Toolchain, Designer, localization, XML documentation and source conventions |
| [Testing](../tests/README.md) | Managed, native-host, provider, JavaScript and renderer checks |
| [VBA testing contract](vba-testing-design.md) | Callback, instrumentation and native acceptance invariants |
| [Tool reference](reference/vbe-tools.md) | Discovery/invocation and authoritative schemas |
| [Recorded validation](test-coverage.md) | Candidate identity, executed results and measurement scope |
| [Release qualification](release-qualification.md) | Finding register and remaining release boundaries |
| [Updates](updates.md) | Installer builds, updater behavior and distribution contract |
| [Code signing](code-signing.md) | Unsigned 1.0.0 and trusted-signature onboarding |
| [Roadmap](roadmap.md) | Planned work and unresolved scope |

## Repository policies

[Contributing](../CONTRIBUTING.md) · [Security](../SECURITY.md) ·
[Support](../SUPPORT.md) · [Conduct](../CODE_OF_CONDUCT.md) ·
[Third-party notices](../THIRD_PARTY_NOTICES.md) · [Changelog](../CHANGELOG.md)

## Documentation conventions

Guides describe implemented behavior. Recorded validation describes observed
results; a target or architecture description is not a successful test.
Test counts and coverage belong only in the validation page, tied to their source.
Superseded investigations remain in Git history. Test fixture Markdown and
third-party/reference data retain their independent purposes.
