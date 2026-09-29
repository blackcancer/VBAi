# Native renderer stop matrix

Prepared before implementation/testing. No Office process, native DLL loading, HWND, registry or filesystem payload mutation.

| Scenario | Stop result | Managed state | Resources/call proof |
| --- | --- | --- | --- |
| No stop export and inactive | true | inactive | no native call |
| No stop export but active | false | active preserved | module/delegates preserved |
| Native success after active | true | inactive | one native call; pinned module/export retained |
| Native success already inactive | true | inactive | repeated stop remains honest/idempotent |
| Invalid thread 1444 | false | active preserved | one call; module/export retained; owning VBE thread diagnostic |
| Generic native error | false | previous state preserved | no retry/unload |
| Error followed by successful owner retry | false then true | active then inactive | exactly two explicit calls |
| Managed invocation exception | false | previous state preserved | error logged; no retry/unload |
| Error while inactive | false | inactive preserved | cleanup still not confirmed |

The fixture owns a scoped snapshot of renderer static fields and restores it in Dispose; tests are DoNotParallelize. Native stop delegates are fake managed callbacks. Real native hook restoration remains a separate runtime qualification gate.