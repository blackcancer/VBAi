---
name: Bug report
about: Report a reproducible VBAi problem with a sanitized example.
title: "[Bug] "
---

# Bug report

Do not disclose vulnerabilities or confidential project content here. Follow
[SECURITY.md](https://github.com/blackcancer/VBAi/blob/main/SECURITY.md) for private security reports.

## What happened?

Describe the observed behavior and its impact.

## Expected behavior

Describe what should have happened instead.

## Steps to reproduce

1. Start with a disposable project.
2. Describe the exact commands and approval mode.
3. Include the smallest sanitized example that reproduces the problem.

## Environment

- VBAi version or commit and loaded assembly path, with personal paths redacted:
- Host application, version/build and process architecture:
- Windows version, UI language and display scaling:
- Provider/model or CLI version, when relevant (never include keys):
- Conversation mode and editing approval policy:

## Persistence and recovery

Was the change only visible in the editor, synchronized into VBE, saved by the
native application, or verified after reopening? Did cancellation or undo help?
Were other projects open? Do not retry destructive actions merely to answer this.

## Diagnostics

Attach only reviewed, sanitized logs, a local report or screenshots. Remove private
code, document names, usernames, credentials and repository URLs as appropriate.
State which checks were actually run; leave unknown information marked unknown.
