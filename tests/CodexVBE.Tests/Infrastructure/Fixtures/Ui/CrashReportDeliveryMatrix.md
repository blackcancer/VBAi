# CrashReportDelivery matrix

Prepared before tests. Native boundaries are supplied with deterministic settings, credential, HTTP, profile-count and automation adapters. Tests invoke the original orchestration; no actual settings reads, registry profile scans, Outlook/GCM activation, publication or drafts.

| Area | Cases |
| --- | --- |
| Send policy | expected issue URL; null/malformed/http/foreign host/path; GitHub status below400/4xx/5xx; credential refusal; cancellation before fallback; mail uncertain |
| Publish native | settings failure; credential success/timeout/explicit cancellation/error; issue object/null; real GitHubApi with local HTTP handler; full repository/title/body serialization |
| Profile gate | three versions scanned in order; null/zero/positive counts; early exit on configured profile |
| Outlook | active object success/refusal; no profile/type absent/creation failure; no accounts; properties fail before Send; Send success/uncertain; exact To/Subject/Body |
| Cleanup | reverse mail/accounts/session/application order; null/non-COM/COM references; release refusal logged while remaining resources still released |
