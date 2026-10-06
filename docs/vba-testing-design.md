# VBA test explorer implementation contract

[Documentation](README.md)

This page records architectural contracts and acceptance boundaries. The
[usage guide](vba-testing.md) is authoritative for declarations, UI actions,
reports and storage. [Recorded validation](test-coverage.md) is authoritative for
executed checks, loaded assemblies and host qualification. Implementation exists;
production acceptance remains pending and must not be inferred from this design.

## Scope

Discover explicitly annotated tests, run a single test or selected batch, retain
verified outcomes and navigate to original source. The shared VBE layer is the
foundation. Host-specific invocation and document-copy operations are adapters,
not permission to advertise an entire host as qualified.

The requested production scope includes Excel, Word, PowerPoint, Access,
Publisher, Outlook and SOLIDWORKS. Visio and Project are outside that qualification
request even though their process names are recognized by the shared native
capability code. Each host needs separate installed-in-process execution evidence.

Keep .NET Framework 4.8, C# 7.3, x64 and existing add-in identity. This work does not
add an installer, automatically save original documents, change trust policies,
execute unrelated production macros or publish project data.
VBE automation must not use SendKeys or global keyboard shortcuts.

## Implemented responsibilities

| Source | Responsibility |
| --- | --- |
| `VbaTestDiscovery` and `VbaTestModels` | Plain-source annotation discovery, test/fixture identities and blocking diagnostics. |
| `VbaTestRuntimeSource` | Versioned project-local assertions and reviewed dispatch source. |
| `VbaTestRunner` | Serial module/test fixtures, outcomes, cleanup and cooperative cancellation. |
| `VbeTestExplorerService` | Owning-thread project snapshots, revision/capability guards, support installation and shared run history. |
| `VbaTestRuntime` and the result sink/native execution adapter | Registered late-bound COM callback for an already authorized pending attempt. |
| `VbaCoverageInstrumentation` | Procedure-entry probes, explicit production denominator, original-source mapping and diagnostics. |
| `VbaTestCoverageClone` and the coverage service partial | Excel, Word and PowerPoint document-copy preparation and measurement without instrumenting original source. |
| `TestExplorerWindow` and its Designer | Shared-theme UI, filtered selection, grouping, progress, reports, explicit measurement review and idle freshness checks. |
| `VbaTestReports` | Human and compact representations of the same canonical results and coverage report. |
| `LlmVbeTools.Testing` | Nine project-scoped tools behind the existing catalog and conversation guards. |

These responsibilities describe source implementation. A pure test or simulated
adapter does not prove a native invocation, clone, docked window or installed COM
activation works in a host.

## Declaration and discovery boundaries

The implemented grammar uses full-line apostrophe annotations in standard
modules: `@TestModule`, `@TestMethod`, repeatable quoted `@TestCategory`, quoted
`@Ignore`, and the four initialization/cleanup markers. Only explicit Public
parameterless Subs and Boolean Functions are test entry points. Fixture entry
points are Public parameterless Subs. Conflicting fixtures, duplicate annotations,
conditional entry points, unsupported signatures and component types stay visible
as diagnostics. Discovery neither executes VBA nor installs support.

Module names are suggestions, not test authorization. Comments inside strings,
procedure bodies and commented-out declarations do not opt in code. Physical
source locations must remain correct across continuations and statement boundaries.
Classes, properties and production functions with arguments can be exercised from
test bodies without becoming test entry points themselves.

The interaction model follows Microsoft's
[Test Explorer](https://learn.microsoft.com/en-us/visualstudio/test/run-unit-tests-with-test-explorer?view=vs-2022).
[Rubberduck's unit-testing conventions](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing)
inform familiar annotation names; they do not establish drop-in compatibility
with Rubberduck assertion objects, private fixtures or fakes. No third-party
framework source is incorporated by this contract.

## Runtime result authority

Installation of `VBAiTestSupport` is a reviewed source edit with collision checks,
revision revalidation and preserved original/proposed source. Opening the explorer,
discovering tests and starting an ordinary run do not inject a new module.
Recognized framework versions can be upgraded only through the reviewed path;
an occupied name containing unrelated user code is refused.

Support version 3 keeps the public assertion and dispatch entry points stable.
Generated private dispatch leaves contain at most 128 cases and 16 Ki characters
of branch source; routes contain at most 32 children and return immediately after
one match. VBA's non-short-circuit `Or` is never used to invoke alternative
children. These source budgets prevent an unbounded central procedure, but do
not guarantee that every catalogue fits VBA's compiled procedure/module limits.
Updating an earlier owned support module still requires review and a backup.

The late-bound `VBAi.TestRuntime` class binds on the owning VBA thread to an
already pending authorized attempt. The generated support source requests that
attempt using its version and dispatch signature, then publishes its verdict.
The result sink verifies the runtime binding, nonce, source revision, run ID and
test ID; it refuses duplicate, unrelated or stale publication. The callback class
does not authorize or schedule arbitrary tests. Registration and native host
activation remain separate qualification gates.

The registered callback is the default transport for ordinary native runs.
Excel, Word and PowerPoint returned-value adapters provide a fallback when that channel
is unavailable and supply results for copy-based coverage. PowerPoint invocation
uses its PIA `Run` contract with an argument array passed by reference. Word
activates the exact owned document, revalidates identity and the active context,
and uses document-qualified `Application.Run` with zero or two positional
arguments. Activation/completion failures are uncertain and are not retried.
An Immediate echo,
delivery acknowledgement, native menu action or invocation start is never a test
verdict. A wrapper must complete with an accepted outcome. Missing completion or
uncertain effects stop the batch and latch execution unavailable until the host
is inspected and the session is reconciled; no automatic retry is allowed.

Boolean True passes and False fails. Project-local assertions use strict scalar
types and binary string comparison, reject unsupported values and retain failure
even if VBA swallows an assertion error. VBA errors remain distinct from assertion
failure and preserve bounded error details. A normal empty Sub can pass;
unfinished templates should call Inconclusive.

## Identity, scheduling and freshness

Keep COM/VBE, WinForms and native dispatch on their owning thread. Immutable
plain-source snapshots may be analyzed separately; live COM objects must not be
passed to worker threads. Identity includes the exact live project and host path
where available, not only its VBA name. The project revision covers module source
and captured reference metadata. It does not reproduce workbook data, a database
or other external dependencies.

Only one run may be active in a VBE session. Freeze explicit project/test IDs
before dispatch, deduplicate tests displayed in multiple category groups, and
retain identity, revision, protection, design mode and originating conversation
authorization checks before deferred native calls. Changed or closed projects
refuse subsequent work.

Module initialization/cleanup surrounds its selected tests. Test initialization
and cleanup surround each selected body. Ordinary failures attempt applicable
cleanup; uncertain completion or lost authorization refuses further dispatch.
Cleanup failure is retained as an error, not a successful rollback of host state.

The visible idle explorer polls source every two seconds. Changed revisions make
historical outcomes stale without rewriting their recorded verdicts; unreadable
projects lose executable actions. Polling is suspended during execution and
coverage review. It never starts tests. Source navigation and rerunning failed
tests remain explicit operations.

Cancellation stops further scheduling after accepted current work and cleanup.
Closing a running explorer requests cooperative stop and waits for the current
result; it retains its result callback. A blocking COM call, modal VBA dialog or
infinite loop can prevent responsiveness. `End`, IDE reset, host closure and
breakpoints can bypass normal publication. Never reset the VBE, terminate a host
or resend a call to make an uncertain attempt appear complete.

## Coverage metric and preservation

The implemented collector measures **procedure entries**, not statements or
branches. A complete report with a positive known denominator computes:

```text
coverage = distinct entered eligible production procedures
           / eligible production procedures * 100
```

Pass rate is a separate ratio: Passed / (Passed + Failed + Error). Discovered test
count, assertion count and add-in .NET coverage cannot substitute for measured
VBA procedure coverage. Missing or unknown evidence is never 0% or 100%.

Measurement is an explicit Excel, Word or PowerPoint document-copy operation. A saved
supported document is copied, opened as a distinct project,
checked against the original snapshot and instrumented only in that copy. Excel
copy opening and closing temporarily suppress workbook events and restore the
previous setting. PowerPoint creates a macro-enabled `.pptm` copy with format 25;
it has no equivalent `EnableEvents` boundary. Application-level open and
before-close handlers may execute and can prevent normal copy closure.
Before closing its exact owned copy, the PowerPoint adapter sets `Saved` to
`msoTrue`, verifies it and revalidates copy identity. This discards instrumentation
without saving; the original presentation's `Saved` flag is never changed.
Word copies the saved DOCM, DOTM, DOC or DOT file rather than saving the original.
Full source/reference matching refuses unsaved VBA edits that the disk copy does
not contain. `Document.Saved=true` is required before copying and checked again
after copying and opening, so unsaved document content is refused. Word AutoOpen, document
open/close and application handlers and the Normal template are not disabled by
the adapter;
neither trust policies nor an invented Word `EnableEvents` setting are changed.
The original document is neither instrumented nor saved by the operation. The
copy runs in the same host application with its privileges; neither tests nor
event handlers are isolated from external systems or other open documents.

The native compile boundary selects the owned coverage runtime pane and validates
the built-in command (ID 578, button type, no custom action). It executes at most
once, yields to the owning UI and observes the command's disabled state for up to
three seconds. Original/copy revisions, permissions, cancellation, design mode
and exact active-project identity are checked before observation and again after
the successful COM read. A still-enabled control refuses measurement; it does
not authorize a second compilation or any coverage test dispatch.

Tests, fixtures and framework support are intentional exclusions. Eligible
production procedures can be in standard, class, UserForm or document modules.
Every probe retains the original module, procedure kind, line and column; the
plan also retains module source hashes. Conditional declarations can make the
denominator unknown. Unsafe inline
headers, reserved runtime identifiers, capacity limits and unsupported syntax
block instrumentation instead of silently shrinking the denominator.

Reports retain mapped hits, intentional exclusions, diagnostics, original revision,
availability and completion. An incomplete run can retain partial hit evidence but
does not claim a complete project percentage. Coverage preview capability is not
measured availability. Statement mapping, branch probes and broader host-copy
adapters remain extension work, not implemented claims.

The coverage plan includes original and instrumented full source. Copies and plan
files remain under `%LOCALAPPDATA%\VBAi\CoverageRuns`; there is no automatic cleanup,
encryption or upload. Support backups and exported reports have separate privacy
boundaries described in [privacy](privacy.md).

## Required production qualification

Acceptance requires operation-specific evidence against the actual loaded build:

- Native callback activation, owning-thread dispatch and result binding in each
  requested host, including failure, missing completion and cancellation cases.
- Installed explorer opening/docking, theme, accessibility, source navigation,
  filtered single/batch execution and external-edit freshness.
- Excel, Word and PowerPoint copy identity, preserved original live/disk source,
  original-source probe mapping and an independently calculated denominator.
- Excel event-setting recovery, Word saved-source matching, owned activation,
  AutoOpen/document/application/Normal-template effects, and PowerPoint
  application-level open/before-close effects; cancelled closure and retained-copy recovery without retrying an
  uncertain native operation.
- Refusal of unsafe coverage syntax, unknown denominators, collisions, stale
  plans and original-project instrumentation; honest partial/uncertain reporting.
- Identical human/compact canonical metrics and preserved approval/privacy guards
  across direct tools, catalog gateways and deferred dispatch.

Use only owned disposable fixtures. SOLIDWORKS must be preloaded and explicitly
selected; do not launch or terminate it autonomously. Follow the
[host-test opt-ins](../tests/README.md#native-host-tests-are-opt-in). Record results
once in [validation](test-coverage.md), with versions, loaded assembly and skipped
scenarios. Until these gates pass, production completion is unproven.

Persisted playlists, durable run history, broader coverage adapters and
statement/branch coverage remain future work. They must not be described as
available simply because their data or UI contracts are outlined here.
