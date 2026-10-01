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
| `testing` | VBA test discovery, reviewed support installation, guarded batches, results, stop, source navigation, procedure coverage and explorer opening. |
| `git` | Bound-repository inspection, commits, checkpoints, branches and merges. |
| `environment` | Projects, references, editor options and other VBE services. |
| `all` | All schemas permitted by the current mode, not additional permissions. |

The `debug` family includes native Locals, Watches and Immediate inspection,
optional Call Stack capture, pane opening, breakpoint and stepping commands,
watch editing, Quick Watch, and bounded scalar-local inspection. These are
different operations with different effects: expression evaluation and Immediate
execution can run VBA; `inspect_local_scalars` supports only declared simple
scalars/parameters in a paused procedure and does not return a complete Locals
snapshot. Shared VBE context permission and Automatic editing policy are required
for the more sensitive reads or evaluations specified by each schema. A breakpoint
toggle response is not a complete, verified breakpoint inventory.

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

The current testing family is defined in
[LlmVbeTools.Testing](../../src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs):

| Tool | Operation boundary |
| --- | --- |
| `discover_vba_tests` | Inspect annotated tests/fixtures and the current project revision without execution. |
| `preview_vba_test_support` | Return exact generated support source for review without installing it. |
| `install_vba_test_support` | Apply only that reviewed source with revision, mode and approval guards; retain recovery evidence. |
| `run_vba_tests` | Run an explicit project/test selection; `Action="coverage"` explicitly selects instrumented Excel/Word/PowerPoint document-copy measurement. |
| `vba_test_run_status` | Read the same run by Query in human or compact form; do not execute again. |
| `stop_vba_tests` | Request cooperative stop of that exact run; do not reset or terminate its host. |
| `navigate_vba_test` | Open one discovered test after validating its current revision. |
| `vba_test_coverage` | Preview procedures, exclusions and capability; Query reads measured coverage of the exact run. |
| `show_vba_test_explorer` | Open the owned explorer on the exact authorized project, without running tests. |

Obtain exact fields/types through live discovery. Use discovery's
`ExpectedProjectVersion`, explicit test IDs in Items and `ExpectedMode=2` for a
run or support mutation. Previewed coverage capability is not measured
availability; a complete result reports actual entered/eligible **procedures**,
original-source locations, exclusions and diagnostics. Statements and branches
are not measured. Run status and both report formats use the same canonical data.

Coverage creates and retains private document/source artifacts, does not
instrument or save the original document, and cannot isolate external effects.
Copies run in the host application with its privileges; PowerPoint application-level
open/before-close handlers may execute and interfere with closure. Word copies
the saved file and refuses source/reference mismatches before instrumentation;
AutoOpen, document/application handlers and the Normal template are not disabled
by the adapter and remain subject to existing host policies.
Its returned-value path requires exact owned activation and document-qualified
invocation, with no retry after uncertain completion. Trust policies are preserved.
Inspect its preview and concrete execution scope before approval. The same
project, privacy, mode, approval and revision checks remain in both direct and
catalog invocation, including deferred native calls. See the
[VBA test guide](../vba-testing.md) for review, storage and qualification limits.

The main, Editor, Git, Monaco and Testing partial definitions describe required/optional
fields and types. The progressive catalog and privacy checks determine which
schemas and operations are available in a conversation. Validate changes with
the catalog/permission tests rather than updating a disconnected static count.

For a new tool, classify data scope explicitly. Tools without a `Project` argument
may expose global context and must not become implicitly trusted. Add malformed
argument, stale revision, wrong project, denied permission, interrupted execution
and recovery cases where applicable.

See [architecture](../architecture.md), [conversations](../chat-ui.md),
[privacy](../privacy.md) and [testing](../../tests/README.md).
