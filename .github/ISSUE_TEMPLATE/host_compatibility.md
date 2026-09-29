---
name: Host compatibility report
about: Share a scoped success or failure in an application that embeds the VBE.
title: "[Host] "
---

# Host compatibility report

A report about one operation helps more than a blanket "supported/not supported"
label. Use a disposable project and do not change organization security policies
just to run a test.

## Environment

- Application and exact version/build:
- VBE/process architecture:
- Windows, language and display scaling:
- VBAi version or commit and loaded build:
- Relevant trust settings, with no sensitive values:

## Operations observed

Mark each relevant item **passed**, **failed** or **not run** and give the precise
scope: add-in loading, code read/edit, references, UserForms, Monaco, compilation,
execution/debugging, native Save, explicit document-save adapter, Git or appearance.

## Reproduction and evidence

Provide sanitized steps, a minimal fixture and exact errors. Identify the save
path used and whether the result survived closing/reopening the document. A native
application Save does not by itself qualify the `save_host_document` adapter.

## Recovery and isolation

Were original settings restored and unrelated projects left unchanged? Record any
uncertain outcome; do not claim a pass based only on the absence of an error.
