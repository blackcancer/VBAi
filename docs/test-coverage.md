# Recorded validation

## Latest recorded project checks

At source **`fea0184`** on **2026-09-29**, the merged Debug/x64 solution built with
no errors or warnings. The full VSTest run recorded **2,004 passed, 0 failed and
28 conditional skips** in `artifacts/branch-integration/final-tests/global-final.trx`.
The accepted run followed corrected asynchronous-tool test expectations and
translations for three notifications in the initial pass; the initial pass is
not the green result. The add-in DLL in
`artifacts/branch-integration/final-build/VBAi/Debug/net48/` was built in an
isolated output and was not installed for that run.

The same integration validated **46 WinForms Designer surfaces**, **27 Designer
metadata checks** and **17 Monaco distribution files** against their regenerated
versions. The mirror inventory was **253 test mirrors for 312 production sources**.
These checks do not establish native host behavior or current instrumented code
coverage. No new C# coverage collection was recorded for this integration.

## Earlier and host-specific evidence

The following runs use different builds and must be read separately:

| Evidence set | Recorded result | Scope and limitation |
| --- | --- | --- |
| Post-rename global VSTest run, `5c860a3` | 1,876 passed, 0 failed, 21 skipped. | Precedes the responsiveness merge; no new instrumented coverage collection. |
| Post-rename Designer checks | 46 surfaces validated. | Loading/editing/resizing/serialization is not full native UI qualification. |
| Office host batch | 4 passed, 0 failed, 1 skipped. | Seven scenario groups per successful host: Word, PowerPoint, Access and Publisher. Outlook blocked by missing setup/profile. |
| Post-rename Excel startup | Loaded DLL identity, automatic Monaco startup, native code behind Monaco, resize and Object Browser layout checked. | No macro execution; not a new qualification of every editor or other-host feature. |
| Later Excel audit checkpoint, candidate from `d75d761` | Native breakpoint, isolated workbook persistence and bounded local-scalar inspection each passed after fixture follow-up. | First scalar attempt was refused because its code pane was not active; explicit navigation before Run Sub passed. Direct cold-navigation execution, save/reopen and other hosts were not qualified by these cases. |
| Historical instrumented run, source `953c84f` | 30,521/30,521 lines and 31,195/31,195 branches; 1,867 passed, 0 failed, 21 skipped. | Predates PR #13 and the later rename changes. **Not current coverage.** |

**Current instrumented coverage after the latest changes is not established by
these records.** Do not carry the historical 100% into a README badge or a current
release claim.

The Office batch verifies common VBE behavior and helper-driven save/reopen. It
does not qualify VBAi's own document-save adapter in every host. The
[compatibility table](compatibility.md) distinguishes recorded refusals, missing
adapters and the historical PowerPoint HWND failure from its later unqualified
code correction. A skipped Outlook test is not a pass.

## Provenance

The pre-refactor records are preserved at the immutable baseline:

- [Merged integration record](https://github.com/blackcancer/VBAi/blob/101a7cd/docs/test-coverage.md).
- [Office and post-rename history](https://github.com/blackcancer/VBAi/blob/99b5f25b39e32131b4c8486ff2d1ae467206b39e/docs/test-coverage.md).
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
