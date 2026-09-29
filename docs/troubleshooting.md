# Troubleshooting

Start with the VBAi build/commit, host and VBE versions, process architecture,
Windows version and the exact failing operation. The About window can copy
technical details without reading project code or provider credentials. Review
all diagnostics before sharing them.

## The add-in does not load

Confirm a compatible **64-bit** VBE process and the required runtime dependencies.
Close hosts before rebuilding the installed output. Run
`tools/Test-VBAiInstallation.ps1` and check the type library, ProgIDs, CLSID,
CodeBase and VBE add-in registration using [source-build setup](installation.md).

A visible menu or registry entry is not enough: query `status` on the intended
`VBAi.<PID>` bridge and compare the loaded assembly/build identity. A hidden VBE
window may leave the host and CLR running. Restart the entire application after
replacing the DLL. In sandboxed development shells, verify the real registry,
not only an isolated view.

## The chat is misplaced or cannot dock

The VBE stores native tool-window layout. A new profile may need manual placement.
If the available pane is smaller than the chat minimum, VBAi can fall back to a
usable floating native pane rather than forcing an unusable layout. Check the
`VBAi.ChatToolWindow` registration if tool-window creation fails.

Report DPI, monitor layout, UI language and the native window involved. Do not
reset all application preferences or delete conversation data to fix positioning.

## A provider cannot connect or no models appear

Verify the selected provider, full endpoint, credentials and account entitlement.
A stored key takes precedence over an environment fallback. Restart the host after
changing environment variables. For Codex/Copilot, use the provider's VBAi-isolated
sign-in and a native executable, not a `.cmd` launcher.

An HTTP 401/403, quota refusal, missing deployment or unsupported tool-calling model
is not the same failure. Test a small non-sensitive request. Do not post API keys,
authentication files or raw provider responses containing private context. See
[providers](providers.md).

## An action is refused or a turn pauses

Check the selected project, Discussion/Plan/Agent mode, approval policy, project
access and shared-context consent. A stale hash or changed mode requires a fresh
read. Some native operations require Automatic approval; investigate the action
before deliberately changing policy.

A paused HTTP tool workflow can be resumed with the same saved profile. An
interruption with no terminal result is uncertain. Inspect the live state before
repeating an edit, import or execution. [Conversations](chat-ui.md) explains pauses
and queued-message handling.

## Changes are visible but not saved

Background synchronization only updates live VBA. Ctrl+S requests a native save
after synchronization; an adapter save is a distinct route. Resolve native dialogs
and inspect the result. An unavailable save adapter does not imply that the host's
own Save command is unavailable.

Save through the host and verify by reopening a disposable file. Keep recoverable
exports and drafts until the result is confirmed. Do not clear a conflict by
blindly replacing the live module. See [editor recovery](modern-editor.md) and the
[host-specific save results](compatibility.md).

## Git imports or recovery are blocked

Inspect local changes, divergence, manifest validity and reference compatibility.
Do not delete the Git directory: it can contain unpushed commits, checkpoints and
the only pre-import backup. A failed partial import must be reconciled before
another synchronization. Keep `.frm`/`.frx` pairs together and do not create missing
host objects by importing a document module. See [Git and GitHub](github-integration.md).

## Appearance and native palette recovery

The add-in/Monaco theme and the experimental native VBE dark theme are distinct.
Native theming depends on VBE/Windows surfaces and can vary by host, DPI and language.
Disable the experimental option when isolating an appearance issue.

Do not delete a palette recovery file to dismiss a warning. The theme reconciles
colors changed after application, archives the prior recovery state before replacing
it, and removes active recovery only after a verified restoration. Invalid files
and incompatible categories can still be refused. Preserve these files when
reporting a recovery failure.

Avoid concurrent native palette experiments in different hosts: preferences can
be shared. Forced termination is not a normal theme-reset procedure.

## Report a problem safely

The in-app problem window previews a report and saves a local Markdown copy under
`%LOCALAPPDATA%\VBAi\CrashReports` before transmission. GitHub delivery targets the
VBAi repository using the selected GitHub account, not the repository bound to a macro.

If credentials are unavailable or an eligible request is refused, delivery can
fall back to the default account in configured **classic Outlook**, addressed to
`init-sys-rev@hotmail.com`. This can send mail; it is not always only a draft. If
classic Outlook is unavailable, the default mail client gets a draft containing
the saved-report path; attach the report manually and send it yourself.

Timeouts, interrupted delivery or ambiguous responses can produce an uncertain
state and block repeat submission. Check GitHub and the mail client before trying
again. Automatic managed-error capture saves a pending report locally; it does
not send during fatal termination and cannot capture every native host crash.

Reports exclude raw exception messages, logs, settings, keys, conversations and
VBA by default. Your own description can still disclose private information.
For a vulnerability, use [private security reporting](../SECURITY.md) instead.
