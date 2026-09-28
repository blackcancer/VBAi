# CrashReportWindow matrix

Prepared before tests. Delivery, clipboard, links and storage are injected; no real reports, browser, email or GitHub publication.

| Area | Cases |
| --- | --- |
| Configuration/preview | null report, empty title, oversized description, report absent during designer initialization |
| Appearance | light/dark/high contrast, disposed, foreign thread marshaled to owned STA, designer/runtime disposal |
| Send | busy/submitted/uncertain repeat refusal; invalid body; save refusal; GitHub success; async success after disposal; async error after cancellation/disposal; transport error with live form |
| Email | busy/submitted/uncertain repeat refusal; save refusal; Outlook/draft success; uncertain Send and arbitrary failure |
| Other actions | copy success/failure, save success/failure, issue open success/failure via injected link boundary |
| Closing/modal | busy user close canceled, other reasons allowed, idle close allowed, absent/malformed owner, supplied native owner |
