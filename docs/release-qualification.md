# Release qualification

[Documentation](README.md)

**Release-wide acceptance is not established.** The selected operation matrices
below have scoped acceptance on their recorded binaries. Combining results from
different candidates does not qualify the latest integrated build.

## Current acceptance summary

| Scope | Decision | Candidate/evidence |
| --- | --- | --- |
| Ollama and embedded Office assistant (Q028) | Accepted for the selected CPU/model/sampling profile and six Office hosts | `85486c00`; [recorded result](test-coverage.md#qualification-register) |
| Word Git and Excel UserForms (Q024/Q027) | Accepted for the complete prepared operation matrix on the real desktop | `c593d6af`; [recorded result](test-coverage.md#qualification-register) |
| Access/Publisher adapters (Q012) | Accepted existing-document persistence and explicit metadata contract | Frozen `2106fd95` product; [recorded result](test-coverage.md#qualification-register) |
| SOLIDWORKS core/editor/local assistant (Q014) | Accepted selected 2019/2025 workflows on the recorded private desktop | Frozen `ddf638b2` product; [recorded result](test-coverage.md#qualification-register) |
| SOLIDWORKS native creation/publication/reopen (Q020/Q030) | Accepted selected 2019/2025 main-desktop workflows; unsafe generic Open refused | Frozen `08689325` product; [recorded result](test-coverage.md#qualification-register) |
| Excel options (Q026) | Accepted selected current-behavior matrix under the maintainer's criterion | `08d7420`; [native boundaries](test-coverage.md#qualification-register) |

[Recorded validation](test-coverage.md) summarizes all qualification findings and
coverage measurements. Candidate hashes and detailed receipts remain in the
evidence indexes, local campaign proof sets and Git history. The original complete-empty Ollama cause,
historical break-mode crashes and other excluded behavior remain unresolved;
scoped acceptance does not claim their causes have been repaired.

## Release criteria

Before claiming a release, identify one integrated candidate and verify applicable
managed, JavaScript, native-renderer, Designer, authenticated-provider and native
host scopes. Record skipped, blocked and unsupported operations explicitly.
The complete managed line/branch target still requires a current measurement.

Native acceptance requires loaded assembly identity, exact disposable project,
operation oracles, original normal process exits, restoration and persistence
readback where relevant. Refusing an unsupported operation can pass a refusal
scenario without providing that capability. Unknown outcomes prohibit retries
of native mutations. Trust policies and existing production documents are preserved.

The agreed host scope includes installed Microsoft 365 x64, classic Outlook and
the independently selected SOLIDWORKS versions. Visio, Project, x86 and other
Office/application builds have no inferred acceptance.

## Finding register

The complete Q001–Q030 register, including earlier results and remaining limits,
is maintained in [test coverage and qualifications](test-coverage.md#qualification-register).

## Review and publication

Use [the testing guide](../tests/README.md) to select the applicable scope.
Publish the candidate, test changes and maintained documentation together; do not
replace native evidence with detached screenshots or protocol mocks.
Distribution must include the declared [licenses](../LICENSING.md) and third-party
notices. The signed installer/release-payload milestone remains later work, as
described in [updates](updates.md).

Superseded checkpoints and investigations remain in Git history. Local artifacts
are retained evidence, not a public package or a license grant.
