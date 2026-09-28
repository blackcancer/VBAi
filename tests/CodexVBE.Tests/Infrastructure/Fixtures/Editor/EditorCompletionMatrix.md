# Editor completion matrix

| Family | Whole batch before execution |
| --- | --- |
| Tools | every initial/post-capture busy, closing, disposed, readiness guard; reference/native component identity; missing document; snapshot null/stale/current; every selection bound; version validity; read-only/divergent native; failed/stale/newer apply; plan changed; native-write failure; post-write newer typing; synchronize exact SHA before/after worker; callback optional/present; draft preserve/clear |
| Language | missing/null/stale request; managed/native source overlays and foreign projects; broken/valid references; worker existing/new/fault; reply ready/disposed; definitions missing/native/managed and clamped positions |
| Window | initialization and script readiness/lifetime guards; owned native browser callbacks and denied navigation/resource policies; message size/type/version dispatch; document tab/recovery/capture revision; preparation and COM write faults; status generations/contrast/error; all diff/reload/resolve/restore/close guards and catches; asynchronous close wait and resource disposal |
| LanguageIndex | conditional directives and unmatched End; modifiers only; incomplete procedure/property; all source/procedure kinds; types/default Variant; conditional symbols and scoped physical ranges |
| SyncWorker | real worker thread snapshots and queued evaluation; action error; disk save error; adding completed/disposed; double dispose after worker exits |

All tests use owned forms, temporary files, in-memory source modules or synthetic VBIDE contracts. No Office process, credentials, real settings or COM registration.