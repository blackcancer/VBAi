# SendAsync finalization completion matrix

Prepared before the batch. Global inventory has one missing branch: the absent current session in the nullable BudgetPaused guard in SendAsync finally (line 468).

| Current session at provider failure | Expected result |
| --- | --- |
| Absent | Original provider failure appears in transcript; draft restored; no session save or compile; HTTP client and active turn released; busy false |
| Present, paused | Original provider failure and restored draft; budget pause preserved; no compile; same cleanup |
| Present, unpaused | Original provider failure and restored draft; no compile without applied changes; same cleanup |

The fake HTTP handler invokes the state transition synchronously at its response boundary and throws a named IOException. No timers, arbitrary delays, network, real settings, host input or production seams.

## Qualification

96 passed, 0 failed, 0 skipped (31 seconds). ChatWindow.cs and SendAsync reach 100% lines and branches in both JSON and Cobertura. Evidence: artifacts/sf/final/contracts.trx and cee9db28-c472-4ccf-aa84-aff3bf20e763/coverage.{json,cobertura.xml}. Layout: 208 mirrors / 265 source files, PASS. Production unchanged, no exclusions.
