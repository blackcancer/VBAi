# VBA test explorer

This implementation is under review. Registered native execution remains blocked
by code-window focus validation in the latest disposable Excel trial. Production
acceptance is incomplete; see [recorded validation](test-coverage.md).

VBAi discovers explicitly annotated tests in live VBA projects and presents them
in the VBE's **View > VBAi test explorer** window. It uses VBAi's shared theme,
controls, localization and native tool-window container, with a floating-window
fallback if the host cannot create the container.

This feature tests your VBA code. The add-in's own .NET test suite and recorded
qualification are separate: see [testing](../tests/README.md) and
[recorded validation](test-coverage.md).

## Declare tests

Place `'@TestModule` in a standard module's declaration section and
`'@TestMethod` in the comment block above each explicit Public parameterless Sub
or Public parameterless Function returning Boolean. Names alone do not opt in
a module or procedure. Tests can call production functions with arguments,
classes or properties within their bodies.

```vb
Option Explicit
'@TestModule

'@TestMethod
'@TestCategory "Arithmetic"
Public Function AddsTwoValues() As Boolean
    AddsTwoValues = (2 + 3 = 5)
End Function

'@TestMethod
Public Sub ComparesValues()
    ' Arrange and act before asserting.
    VBAiTestSupport.AreEqual CLng(5), CLng(2 + 3), "Unexpected sum."
End Sub

'@TestMethod
'@Ignore "Waiting for a disposable external dependency."
Public Sub DeferredScenario()
    VBAiTestSupport.Inconclusive "Not implemented yet."
End Sub
```

Use `'@ModuleInitialize`, `'@ModuleCleanup`, `'@TestInitialize` and
`'@TestCleanup` on Public parameterless Subs for fixtures. A module can declare
one of each. Module fixtures run once per selected module; test fixtures surround
each selected test. Ordinary failures still attempt cleanup. Uncertain native
completion, stale revisions and lost permissions prevent further dispatch.

Categories are repeatable. Markers are case-insensitive, full-line apostrophe
comments; strings and markers inside procedure bodies do not declare tests.
Private/parameterized/conditional entry points, classes, UserForms, document
modules, `Option Private Module` and malformed annotations display diagnostics.
They are not silently treated as runnable.

The conventions are inspired by [Rubberduck](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing).
Existing Rubberduck assertion objects and private fixtures are not automatically
compatible. Boolean Function tests are a VBAi extension.

## Prepare and run

1. Select the exact project in the explorer. Paths distinguish saved projects
   with the same VBA name. Refresh reads source without executing VBA.
2. Check the displayed execution capability. The shared in-process result channel
   requires the registered x64 `VBAi.TestRuntime` COM class and a live, owned VBE
   thread. This callback is the default transport for ordinary native runs;
   Excel, Word and PowerPoint also have returned-value adapters used as a fallback
   when the callback is unavailable and for copy-based coverage. A capability
   response describes current prerequisites; it does not qualify a host for
   production use. No trust settings are changed. See the boundaries below.
3. Choose **Install test support**, review the diff, and apply the generated
   `VBAiTestSupport` module. This is an explicit source edit, not an automatic
   discovery step. Occupied module names refuse installation unless they contain
   the recognized versioned framework header. Replacements are backed up first.
4. Refresh, select tests or a module/group, and run the selection or visible scope.
   Search and outcome filters restrict that scope. Category groups deduplicate
   tests appearing in several categories. Runs are serial within one VBE session.
5. Inspect the details, open test source, or rerun failed/error tests. While visible
   and idle, the explorer checks source every two seconds. An edit to any project
   module or reference invalidates the previous revision and grays historical
   results; it never runs tests automatically. A closed or unreadable project
   disables its actions. Update the
   reviewed dispatch module if you add, remove or rename callable tests/fixtures.

`VBAiTestSupport` provides IsTrue, IsFalse, AreEqual, Fail and Inconclusive.
AreEqual uses strict matching scalar types and binary string comparison; it does
not coerce a Long into a Double or a string into a number. Unsupported arrays,
objects, Null and Empty are refused. An assertion remains failed even if a test
swallows its error with `On Error Resume Next`. An empty Sub that completes
normally can pass; use Inconclusive for an unfinished test.

A Boolean test passes only on True. Assertions and False produce Failed; other
VBA errors produce Error, preserving their number and message. The verified
wrapper supplies the result through the callback channel or the Excel/Word/PowerPoint
returned-value adapter, rather than an Immediate echo. Callback results are bound to the pending
attempt, support signature, revision, run and test. Missing or rejected completion,
or a native call with uncertain effects, stops further dispatch without retry.
Inspect the host before reconnecting VBAi after such a failure.

**Stop after current test** is cooperative. Queued calls can be refused before
dispatch, and verified current work can finish its cleanup. It does not forcibly
interrupt a modal dialog, infinite loop or blocking host call. VBA `End`, IDE
reset, host closure and breakpoints can bypass normal result publication.

Tests execute with the host's privileges. Use disposable projects for development
qualification; framework cleanup is not a sandbox or a guaranteed rollback of
files, workbook data or external systems.

## Results for people and models

The explorer shows success checkmarks and failure crosses with semantic colors;
labels remain visible and high-contrast mode uses system colors. Details retain
duration, phase, assertion/error messages, source location and tested revision.

Two reports are generated from the same run:

- **Readable report**: localized project/run identification, counts, pass rate,
  per-test outcomes and messages, measured procedure coverage, original-source hit
  locations, exclusions and diagnostics when a coverage run was requested.
- **LLM report (JSON)**: versioned compact JSON with stable test IDs, project,
  revision, counts, outcomes, durations in milliseconds, error numbers and an
  explicit uncertainty flag and the same coverage data. Numeric formatting remains
  culture-independent.

Copy or export the chosen report locally. Results remain in memory for the
session; the service retains at most twenty run records. A project switch clears
the window's report view. Reports for changed source are historical, not current
qualification. The pass rate is Passed / (Passed + Failed + Error); inconclusive,
ignored, cancelled, blocked and uncertain results are separately identified.
An empty denominator is unavailable, not a successful percentage.

## Measured procedure coverage

The implementation can collect **procedure-entry coverage on a separate Excel,
Word or PowerPoint document copy**. These paths still require native qualification; implemented code
is not production acceptance. Statement and branch coverage are unavailable.
Pass rate, discovered test count and .NET coverage are different metrics.

Use **Run selected with coverage** or **Run visible scope with coverage**. The
review identifies the project, original revision, selected tests and retained
copy location. Discovery and ordinary runs never install instrumentation.
The coverage capability preview explains unsupported syntax before execution.

The copy is retained under `%LOCALAPPDATA%\VBAi\CoverageRuns\<unique-run-id>`.
Only the distinct copy is instrumented; its sources must match the original
snapshot before rewriting. The original document is neither saved nor
instrumented by measurement.

| Host adapter | Copy and event boundary |
| --- | --- |
| Excel | `SaveCopyAs` creates a separate `.xlsm`, `.xlsb` or `.xls` workbook. Saving the copy, opening and closing temporarily suppress workbook events and restore the previous setting. |
| Word | Copies the saved `.docm`, `.dotm`, `.doc` or `.dot` file without saving the original. Full source/reference comparison refuses unsaved VBA edits that are absent from the copy. AutoOpen, document open/close handlers, application events and the Normal template are not disabled; no Excel-style `EnableEvents` boundary or trust-policy change is introduced. |
| PowerPoint | `SaveCopyAs` creates `coverage.pptm` using macro-enabled presentation format 25. Saved `.pptm`, `.ppsm`, `.potm`, `.ppt`, `.pps` and `.pot` documents are accepted; presentation add-ins are excluded. PowerPoint has no Excel-style `EnableEvents` boundary: application-level open and before-close handlers may execute. |

Word returned values use the published `_Application.Run` interface with all 30
optional by-reference slots. Only zero or two explicit arguments are accepted;
unused slots contain `Type.Missing`. Word resolves `Module.Procedure` only after
the adapter proves that exactly one loaded VBE project exposes that module and
that its COM identity is the exact owned project. The proof is repeated after
activating the owned document. A collision, unreadable project, changed identity
or ambiguous filename refuses dispatch.

The generated coverage reset and snapshot functions accept two ignored optional
arguments. Word supplies them explicitly to preserve the Boolean and Boolean-array
returns observed at the native boundary; Excel and PowerPoint retain their
zero-argument calls. This adaptation applies only to the generated coverage
functions and never changes a user's procedure signature.

The adapter refuses `AutomationSecurity` set to `ForceDisable` before dispatch or
copy creation; it never changes the host policy. The
[Word signature](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._application.run?view=word-pia)
and [file-opening policy](https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity)
describe these host boundaries.

All three adapters validate separate document, project and path identities. They run
the copy in the same host application with its privileges and environment. This
is not an external-system sandbox: tests and application-level handlers can
affect other documents, files, databases or external systems. A close failure or
cancellation can leave the copy open; inspect uncertain outcomes before acting.
Word's returned-value adapter activates only the exact owned document, revalidates
its project/path, active context and unique module ownership, then invokes the module-qualified
`Application.Run` with zero or two positional arguments. Activation or macro
completion failures are uncertain and never trigger automatic retry. A Word
measurement requires `Document.Saved=true`, checked before copying and again
after copying and opening. Save all document changes before requesting coverage.
Native Word and PowerPoint qualification remains pending.

The metric is distinct entered eligible production procedures divided by eligible
production procedures. Tests, fixtures and framework support are explicitly
excluded. Probe records map back to the original module, procedure kind, line and
column. Standard, class, UserForm and document modules can contribute eligible
production procedures; they do not become test entry points.

Conditional procedure declarations with an unresolved denominator, unsafe inline
headers, reserved runtime identifiers and other blocking syntax diagnostics
refuse instrumentation. These cases are not silently removed to improve a
percentage. The preview distinguishes intentional exclusions, blocking diagnostics
and whether the denominator is known.

A complete measurement with a positive known denominator can show a percentage
and hit/eligible count. Partial measurements retain available hits and diagnostics
without claiming a complete project percentage. Missing, uncertain or unknown
measurements are explicit; they are never converted to 0% or 100%. A later source
revision makes previous measurements historical. Do not merge different revisions
or compare procedure percentages as if they measured statements or branches.

## Assistant operations

Use `discover_tools` with family `testing` to obtain the live schemas from
[LlmVbeTools.Testing](../src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs). They cover test
discovery, support preview/installation, explicit single/batch execution, status,
cooperative stop, navigation, coverage and opening the explorer. The gateway retains the
same project, privacy, mode, approval and revision checks as direct invocation.

Inspect discovery and use its `ExpectedProjectVersion` and exact test IDs for a
run; `ExpectedMode=2` is required. Installation accepts only the exact reviewed
preview Text for that revision. Run status uses the returned Query and an Action
of `human` or `compact`; compact status embeds structured JSON rather than an
escaped report string. Starting a run is not evidence of success. Poll the same
operation; never rerun merely because status is delayed.

The assistant installation gateway accepts reviewed support `Text` up to
1,048,576 characters. The run-selection limit of 10,000 distinct IDs does not
guarantee that every catalogue's generated support fits that separate bound;
long names and fixtures also consume support capacity. An oversized installation
is explicitly refused before dispatch. Pagination does not increase this source
installation limit.

Run status pages both compact and readable details. `Offset` is a zero-based,
nonnegative Int32; `Limit` accepts 1–100, with omitted/zero selecting 100. Compact
reports retain version `v=1` and all existing fields, adding `total`, `offset`,
`limit` and `nextOffset`. Counts, pass rate, coverage percentage, revision and
uncertainty always describe the entire run. Follow `nextOffset` until null: tests,
coverage probes, exclusions and diagnostics each use the same offset, with
separate collection totals. Some pages may therefore contain coverage details
after all test results have been returned. Every returned assertion message is
complete; a short collection page does not imply that every other collection is
finished. Historical status retains its stale-source flag on every page.

The explorer's local readable and compact exports remain complete canonical
reports. The local JSON serializer allows at most 536,870,912 serialized
characters; this is a character limit, not a UTF-8 file-byte guarantee. A larger
report fails explicitly without truncating messages; use paged status to inspect
the retained run. This larger limit is not used for bridge or LLM pages.

`vba_test_coverage` without Query previews eligible procedures, exclusions and copy
capability without editing or executing. Preview availability is not a measured
result. `run_vba_tests` with `Action="coverage"` explicitly selects the copy-based
measurement path; Query on `vba_test_coverage` reads that run's measured report.
Coverage preview and measured Query results accept the same `Offset`/`Limit`.
Their existing public fields remain available; `Probes` (preview), `Hits`
(measured), `Exclusions` and `Diagnostics` contain one page. `Total` is the longest
collection; individual totals and `NextOffset` allow complete retrieval while
eligible/hit counts and percentages stay global. Unsupported plans exceeding the
16,000-probe instrumentation capacity can still be inspected in pages; pagination
does not make those plans executable.
`show_vba_test_explorer` opens the window on the exact authorized project and never
starts a test. A busy explorer cannot switch a frozen run to another project.

Each deferred invocation rechecks the originating conversation's authorization.
Switching mode, revoking access or changing the editing policy can stop the
remaining batch. Shared-context permission is required for execution and stop
tools as described by their schemas. The UI and assistant use the same active
run exclusion, support source and result history.

## Storage and qualification limits

Support installation writes preserved source and proposed source under
`%LOCALAPPDATA%\VBAi\TestSupportBackups\<unique-edit-id>`. Installation does not
save the host document. Backups are retained for manual recovery; there is no
automatic retry/restoration or cleanup of this folder. Exported reports can
contain private paths and messages. No upload is implied by this feature.

Coverage folders retain the host document copy and `coverage-plan.json`. The plan
contains original and instrumented **full source**, hashes and source mapping;
it is private project data, not an anonymous coverage summary. Source changes in
the measurement copy need not be saved when that copy closes. Retained disk
artifacts are not automatically removed, encrypted or uploaded. Keep required
recovery evidence before deleting files manually.

The shared appearance follows the existing theme and supported locale catalogues.
Technical dynamic diagnostics and VBA error messages may remain in their original
language. See [recorded validation](test-coverage.md) for the actual tested build,
native execution boundary, DPI, skipped scenarios and remaining qualification.

The shared callback path recognizes the known x64 VBE host process names in the
source. The requested production qualification scope is Excel, Word, PowerPoint,
Access, Publisher, Outlook and SOLIDWORKS; Visio and Project are excluded from
that request. Recognized names, simulated adapters, passing discovery and a
detached UI render do not qualify native callback execution, installed docking
or coverage. Those gates require owned-host evidence with the actual loaded
assembly and host version. Production completion remains unproven until those
operation-specific gates pass.

## Pending source-navigation integration

Test navigation currently selects the native VBIDE code pane and test line.
The visible Monaco document can remain on the previous module. Opening and
confirming the matching modern document before reporting navigation success is
still pending. The interrupted asynchronous implementation is not included.
