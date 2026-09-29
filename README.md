<p align="center"><img src="assets/icons/assistant.png" width="88" alt="VBAi logo"></p>

# VBAi

**Your AI agent for VBA. A modern workspace inside the Visual Basic Editor.**

VBAi brings AI-assisted development, a modern code editor and Git workflows into
applications that host the **Visual Basic Editor (VBE)**. Work with the VBA project
that is already open: understand existing code, review changes, build UserForms,
and investigate problems without moving between a chatbot and your editor.

The product targets the VBE, not a fixed pair of applications. Its shared VBIDE
integration provides the core experience; host-specific compatibility layers add
operations such as document persistence and procedure execution.

[Get started](docs/getting-started.md) · [Documentation](docs/README.md) ·
[Compatibility](docs/compatibility.md) · [Contributing](CONTRIBUTING.md)

> **Development preview.** The current build targets Windows and 64-bit VBE hosts.
> A standalone installer is planned for a later milestone. Developers and early
> testers can use the [source-build setup](docs/installation.md). Architectural
> compatibility is not a claim that every operation has been tested in every host.

## What you can do

| Capability | In practice |
| --- | --- |
| AI-assisted development | Discuss a project, plan an intervention, or let an agent use controlled tools against live VBA code. |
| Modern editing | Edit in Monaco with completion, signature help, navigation, formatting and native compilation diagnostics. |
| Review and recovery | Inspect code diffs, reject stale edits and undo supported changes without silently replacing newer work. |
| UserForm development | Inspect forms, add controls, adjust layouts and work with events through the native designer. |
| Native debugging | Use compilation, breakpoints and stepping through VBE services, subject to the host's capabilities. |
| Git and GitHub | Version exported VBA sources, compare changes, create checkpoints and use branches and merges. |
| Provider choice | Connect a supported cloud service, a CLI-backed provider or a compatible local model server. |

## A typical workflow

1. Open a saved, backed-up project in its host application's VBE.
2. Start in **Discussion** to understand the code and choose the context to share.
3. Switch to **Agent** when ready, selecting the editing approval policy deliberately.
4. Review the resulting changes, compile, and test the intended behavior safely.
5. Save the document in its host application; commit or publish sources separately.

For example: *“Explain why this procedure fails when the input is empty. Propose a
minimal correction and keep the public interface unchanged.”*

AI output is not a correctness guarantee. Compilation does not establish runtime
correctness, and undoing code does not reverse a macro's external side effects.
See [privacy and safety](docs/privacy.md) before sharing professional or sensitive
projects with a provider.

## Built for the VBE ecosystem

The shared layer works with projects, modules, references, code panes and native
editor services. Application-specific adapters handle operations that VBIDE alone
does not provide. A missing save adapter does not make a host irrelevant to the
project; it limits that operation until an adapter is implemented and qualified.

The [compatibility guide](docs/compatibility.md) separates platform requirements,
implemented adapters and observed results. It includes the current Office and
SOLIDWORKS test coverage without presenting those hosts as an exhaustive list.

## Choose your AI connection

Codex is the default provider. Other integrations include OpenAI API, Claude,
GitHub Copilot, Gemini, Mistral, DeepSeek, OpenRouter, Azure OpenAI, Grok, Groq,
Amazon Bedrock, Ollama, LM Studio and a custom OpenAI-compatible endpoint.

Provider accounts, usage limits and any service charges are separate from VBAi.
A listed integration does not guarantee that every model supports tool calling.
Use the [provider guide](docs/providers.md) for authentication and configuration.

## Project information

- [Development guide](docs/development.md) and [architecture](docs/architecture.md).
- [Testing](tests/README.md), [recorded validation](docs/test-coverage.md),
  [changelog](CHANGELOG.md) and [roadmap](docs/roadmap.md).
- [Help and voluntary support](SUPPORT.md), [security reporting](SECURITY.md)
  and [community conduct](CODE_OF_CONDUCT.md).

### License and third-party components

A license for VBAi itself has not yet been selected in this repository. This
README does not grant a software license. Bundled components retain their own
terms; see [third-party notices](THIRD_PARTY_NOTICES.md).

VBAi is an independent project, not an official product of Microsoft, OpenAI,
GitHub or the vendors of applications that host the VBE.
