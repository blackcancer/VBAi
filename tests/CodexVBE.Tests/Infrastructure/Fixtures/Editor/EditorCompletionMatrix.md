# Editor completion matrix

| Family | Whole batch before execution |
| --- | --- |
| Tools | every initial/post-capture busy, closing, disposed, readiness guard; reference/native component identity; missing document; snapshot null/stale/current; every selection bound; version validity; read-only/divergent native; failed/stale/newer apply; plan changed; native-write failure; post-write newer typing; synchronize exact SHA before/after worker; callback optional/present; draft preserve/clear |
| Language | missing/null/stale request; managed/native source overlays and foreign projects; broken/valid references; worker existing/new/fault; reply ready/disposed; definitions missing/native/managed and clamped positions |
| Window | initialization and script readiness/lifetime guards; owned native browser callbacks and denied navigation/resource policies; message size/type/version dispatch; document tab/recovery/capture revision; preparation and COM write faults; status generations/contrast/error; all diff/reload/resolve/restore/close guards and catches; asynchronous close wait and resource disposal |
| LanguageIndex | conditional directives and unmatched End; modifiers only; incomplete procedure/property; all source/procedure kinds; types/default Variant; conditional symbols and scoped physical ranges |
| SyncWorker | real worker thread snapshots and queued evaluation; action error; disk save error; adding completed/disposed; double dispose after worker exits |

All tests use owned forms, temporary files, in-memory source modules or synthetic VBIDE contracts. No Office process, credentials, real settings or COM registration.
## PR10 Debug and Save matrix (prepared before the batch)

| Area | Complete cases |
| --- | --- |
| Debug observation | absent/busy/managed; null pane; changed design/run/break mode; stable mode; position cached or changed; cached document null/missing/dirty/conflicted/unchanged/revised; execution selected tab present/current/other/new |
| Debug commands | busy wait releases into closing/disposed/normal; initial closing/disposed; null/missing/managed identity; version mismatch; dirty/conflict after synchronization; wrong compile mode; preflight/native/observer failures; clean/dirty/conflicted/changed/foreign diagnostic; every paginated command including breakpoint ID51 and focus; optional browser focus |
| Save | busy wait and initial readiness/closing/disposed/null/missing/managed; project grouping; late closing/disposed; pending dirty/conflict; native save failure; file name failure/null/whitespace/missing/existing; Saved false/true; host false/null/true; warning/error persistence and reset; native focus/identity/missing/disabled/executed command; host status unavailable/available/error |
| PR10 Window | keyboard F9/CtrlS/unhandled modifiers; owned workspace UserClosing; sync/save and valid/invalid assistant action dispatch; selected text/version/native module validation; document status precedence and change resets |