# Tool reference and discovery

This page documents the tool contract, not a second hand-maintained copy of every
schema. The authoritative catalog is constructed by `LlmVbeTools.Definitions` and
its companion definitions in `src/VBAi/Llm/Chat/`.

## Discovery

The model starts with a small core: live status, permitted project/module reads,
Monaco access and catalog gateways. `discover_tools` exposes additional schemas
by family:

| Family | Operations |
| --- | --- |
| `code` | Modules, procedures, navigation, Monaco and guarded refactoring. |
| `forms` | UserForms, controls, containers, layouts and events. |
| `debug` | Compilation, execution and native debugger operations. |
| `git` | Bound-repository inspection, commits, checkpoints, branches and merges. |
| `environment` | Projects, references, editor options and other VBE services. |
| `all` | All schemas permitted by the current mode, not additional permissions. |

For example, `discover_tools` accepts `{"Family":"code"}`. The invocation gateway
uses `ToolName` and `ArgumentsJson`; the latter is the selected tool's serialized
argument object. Invoke only a discovered tool with the exact returned schema.

HTTP conversations add discovered families up to a 64-schema request cap, keeping
the core first. Other functions remain accessible through the gateway. Codex keeps
the initial dynamic core stable and invokes discovered functions through the
gateway. Recursive gateway calls are rejected.

## Contract rules

Names and JSON field names are case-sensitive. Obtain live project identity before
acting; use an unambiguous saved path when names collide. Do not manufacture SHA,
project/tree/window/clipboard versions or Git `ExpectedState` values.

Both direct tools and gateways retain chat-mode, project-access, shared-context,
approval, protection and revision guards. Discovery is descriptive, not an
authorization grant. Discussion/Plan expose inspection rather than unrestricted
execution or mutation.

Some inspection tools open or select native UI or compile code. A pending response
is not proof of completion. An error after a mutation can leave an uncertain or
partial result; inspect it instead of repeating the action blindly.

Use `InvokeAsync` for asynchronous native, Git and Monaco dispatch. The full
catalog is not identical to the bridge's internal command set. A successful tool
response does not automatically establish disk persistence or runtime correctness.

## Source of truth

The main, Editor, Git and Monaco partial definitions describe required/optional
fields and types. The progressive catalog and privacy checks determine which
schemas and operations are available in a conversation. Validate changes with
the catalog/permission tests rather than updating a disconnected static count.

For a new tool, classify data scope explicitly. Tools without a `Project` argument
may expose global context and must not become implicitly trusted. Add malformed
argument, stale revision, wrong project, denied permission, interrupted execution
and recovery cases where applicable.

See [architecture](../architecture.md), [conversations](../chat-ui.md),
[privacy](../privacy.md) and [testing](../../tests/README.md).
