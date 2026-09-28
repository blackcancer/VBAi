# Runtime family coverage matrix

All missing branches in the qualified global JSON were inventoried before the isolated batch.

| Source | Missing paths | Scenarios and observations |
| --- | --- | --- |
| ChatWindow.cs | Activity ownership/disposal callback; null about cleanup; empty send with paused budget; async finally pending exception; constant assistant identity null | Deliver activity to owning live state, reject foreign/disposed state; release partially initialized controls; resume without adding a user question; close during provider failure and assert the recovery exception stack and faulted task; assert the actual embedded identity and about image. |
| ChatWindow.Activities.cs | Legacy entry without Activity; append while running; optional native metadata; long preview; running label and glyph | Upgrade an existing tool stream with no native metadata; preserve appended detail and reported duration; render all five native statuses and assert labels, glyphs, detail and preview bound. |
| CodexAgentActivity.cs | Missing and non-array file changes | Keep the file-change identity and declared declined status, with empty detail for unsupported shapes. |
| CodexAppServerClient.cs | Summary delta without item id; completed reasoning without id; terminal failed and completed running activities | Anonymous summaries do not create activity; terminal failed closes reasoning and commands as failed; terminal completed closes reasoning as completed and unconfirmed commands as interrupted. |
| GitHubApi.cs | Optional issue body null; title/body limits | Accept null body and exact limits of 180/60000; reject oversized title or body before credential resolution and HTTP dispatch; assert serialized payload and returned issue identity. |

No host or network is launched. Provider boundaries use existing owned in-memory handlers and fake App Server transports. No user window, coordinates or shortcut is used.

## Constant resource proof

ChatWindow calls VbeWindowIcons.Icon with the constant assistant name. CodexVBE.csproj embeds assets/icons/assistant.ico unconditionally with logical name CodexVBE.Icons.assistant.ico. A missing input causes build failure. Icon returns null only if the manifest resource is missing; a present but invalid resource throws instead of returning null. Thus every successfully built admissible assembly either produces the icon or throws. The single call site changes identity?.ToBitmap() to identity.ToBitmap(); generic missing-resource guards remain unchanged. The test verifies the actual manifest resource and nonempty image dimensions.

The asynchronous finally exception case asserts failure and disposed state; it does not treat an unverified native result as success.
## Isolated measured result

138 tests passed, zero skipped, with XPlat Code Coverage and no exclusion setting. Each collector used its own absolute BuildOutputRoot; no output path was reused while a collector was active.

| Source | Lines | Branches |
| --- | --- | --- |
| ChatWindow.cs | 390/390 | 288/288 |
| ChatWindow.Activities.cs | 89/89 | 104/104 |
| CodexAgentActivity.cs | 35/35 | 76/76 |
| CodexAppServerClient.cs | 329/329 | 224/224 |
| GitHubApi.cs | 130/130 | 80/80 |

Final report: artifacts/coverage/runtime-families-final/d3b6bb8a-2046-4915-9779-bf0d6fa23bc0/coverage.cobertura.xml.

The absolute build root ends in artifacts/build/runtime-families-final. Filter: ChatWindowStateTests, CodexAgentActivity, CodexAppServerClientTests, GitReviewTests and CrashReportDeliveryTests. The existing report-delivery tests are needed for the structured GitHub failure boundary; they use owned fixtures.

The mirror layout gate passes: 207 dedicated mirrors for 265 production files. This targeted gate does not replace the global suite or real host acceptance.