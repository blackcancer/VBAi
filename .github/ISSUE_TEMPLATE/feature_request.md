---
name: Feature request
about: Describe a VBE workflow and the improvement you need.
title: "[Feature] "
---

# Feature request

## Workflow or problem

What are you trying to accomplish in the VBE? Explain the practical limitation
rather than starting with a prescribed implementation.

## Proposed behavior

Describe a concrete before/after example and how a user would control the action.

## Scope

Is this shared VBE behavior, an AI-provider integration, or an operation that needs
a host-specific adapter? Name an application only where its behavior matters.

## Alternatives and safety

What workaround exists? Does the feature read private data, execute code, modify
a document or affect external systems? Explain the expected approval and recovery.

## Acceptance criteria

What observable result would show that the feature works? Separate native-host
validation from unit tests and avoid assuming that every VBE host behaves identically.
