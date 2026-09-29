# Recorded validation

## Documentation baseline

This summary was consolidated on **2026-09-29** from repository evidence at
**`99b5f25b39e32131b4c8486ff2d1ae467206b39e`**. It reports existing results; the
documentation refactor did not execute Windows/COM, provider or host tests.

| Evidence set | Recorded result | Scope and limitation |
| --- | --- | --- |
| Post-rename global VSTest run | 1,876 passed, 0 failed, 21 skipped. | No new instrumented coverage collection in this run. |
| Post-rename Designer checks | 46 surfaces validated. | Loading/editing/resizing/serialization is not full native UI qualification. |
| Office host batch | 4 passed, 0 failed, 1 skipped. | Seven scenario groups per successful host: Word, PowerPoint, Access and Publisher. Outlook blocked by missing setup/profile. |
| Post-rename Excel startup | Loaded DLL identity, automatic Monaco startup, native code behind Monaco, resize and Object Browser layout checked. | No macro execution; not a new qualification of every editor or other-host feature. |
| Historical instrumented run, source `953c84f` | 30,521/30,521 lines and 31,195/31,195 branches; 1,867 passed, 0 failed, 21 skipped. | Predates PR #13 and the later rename changes. **Not current coverage.** |

**Current instrumented coverage after the latest changes is not established by
these records.** Do not carry the historical 100% into a README badge or a current
release claim.

The Office batch verifies common VBE behavior and helper-driven save/reopen. It
does not qualify VBAi's own document-save adapter in every host. The
[compatibility table](compatibility.md) retains each refusal, missing adapter and
the PowerPoint HWND defect. A skipped Outlook test is not a pass.

## Provenance

The pre-refactor records are preserved at the immutable baseline:

- [Validation history](https://github.com/blackcancer/VBAi/blob/99b5f25b39e32131b4c8486ff2d1ae467206b39e/docs/test-coverage.md).
- [Rename and migration validation](https://github.com/blackcancer/VBAi/blob/99b5f25b39e32131b4c8486ff2d1ae467206b39e/docs/rename-vbai.md).
- [Office scenario details](https://github.com/blackcancer/VBAi/blob/99b5f25b39e32131b4c8486ff2d1ae467206b39e/docs/office-host-qualification.md).

The referenced TRX, coverage and native reports under `artifacts/` were local,
Git-ignored evidence. A path in a historical record is not a downloadable report
in this repository. Relevant recorded locations include
`artifacts/rename-vbai/accepted/global.trx`, `artifacts/rename-vbai/designers/`
and `artifacts/office-hosts/vstest/office-hosts-complete.trx`.

## Measurement boundaries

C# coverage collected in VSTest does not instrument code executing in a separate
Office/SOLIDWORKS process. It also does not measure the native C++ renderer or
JavaScript branch coverage. Simulated provider exchanges are not authenticated
service tests; a synthetic live-provider request is not a VBE workflow test.

Prior SOLIDWORKS debugger checks and live OpenRouter synthetic checks retain their
historical, limited scope. The presence of a tool or test fixture is not evidence
that it ran successfully in every host or account.

## Updating this page

Run the appropriate checks from [testing](../tests/README.md). Replace the current
summary with the tested source/build, date, exact commands, passed/failed/skipped
counts and report location. Report partial or failed runs honestly. Keep old runs
in Git history rather than appending another long chronology here.

No production coverage exclusions may be added merely to reach a target. Keep the
README free of hand-maintained test counters and use published evidence before
adding any CI or coverage badge.
