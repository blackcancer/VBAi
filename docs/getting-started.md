# Getting started

[Documentation](README.md)

This guide is for a first controlled session with VBAi. The standalone installer
will come later; for the current preview, complete the
[source-build setup](installation.md) first.

## Prepare your environment

Use a Windows application with a compatible 64-bit VBE and check
[compatibility](compatibility.md). Save a disposable copy of the document or macro
project. An unsaved project's temporary identity is not a durable conversation
association across restarts.

Open the project's VBE using the host application's normal command. VBAi adds its
assistant and editor entries to the VBE and its configuration to the Tools menu.
Menu wording follows the selected UI language. If the add-in is missing, follow
[troubleshooting](troubleshooting.md) before changing host security settings.

## Configure one provider

Open VBAi settings, choose a provider, complete the required authentication and
save. Choose an available model in the chat. Models that do not support tools may
be unsuitable for agent interventions even when they appear in a catalog.

Codex uses its own CLI-backed sign-in flow. API providers use their configured
credentials; Ollama and LM Studio can use a compatible local server. These are
different access methods, not interchangeable subscriptions. See
[providers](providers.md) for the exact setup and charges outside VBAi.

## Start with an explanation

Select the correct project in the chat. Choose **Discussion** and review the
editing policy rather than assuming that a new installation is read-only.
Use `#` for project/module references, `@` for procedures, or attach an editor
selection. Preview the explicit context before sending.

Try: *“Explain this procedure's inputs, outputs and error handling. Do not change
its code.”*

The selected project is the default data boundary. Other projects and shared VBE
context require separate permissions; [privacy](privacy.md) explains the distinction.

## Apply a small change

Switch to **Agent**, choose an appropriate approval policy and request a narrowly
scoped change. For a first test, **Ask each time** makes approvals visible. Some
native execution tools require **Automatic** and will refuse a stricter policy;
do not relax it merely to get past an unexplained refusal.

Review the diff. A stale revision or a conflict is a reason to inspect the current
code, not to force a replacement. Supported code changes can be undone from the
conversation, subject to conflict checks.

## Verify and save

Compile explicitly or use compilation after an intervention. Test the intended
runtime behavior only on a safe document and when execution is intended.
Background editor synchronization updates live VBA; it is not disk persistence.
Ctrl+S in Monaco synchronizes and requests the native save command for that tab's
project. Resolve any host prompt and check the save outcome.

For an important change, save in the host and reopen the disposable file to confirm
persistence. Git commits, cloud publication and host-document saving are separate
actions. See [the editor](modern-editor.md) and [Git workflows](github-integration.md).
