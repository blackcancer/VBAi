# PR 10 access and budget matrix

| Source | Scenarios | Required observations |
| --- | --- | --- |
| BridgeRequestReader | Invalid budgets, Unicode, LF/CRLF and empty frames, exact byte budget, missing newline, oversized and invalid UTF-8 payloads, deadline closure with successful or faulted outstanding read, healthy local pipe client after rejected peer | Rejected requests never dispatch; timeout closes the owned stream; healthy peers retain service. |
| LlmVbeTools.Catalog | Every classifier alternative, family and all-family discovery, Agent/Discussion/Plan exposure, gateway-only limits, reset and priorities, every malformed gateway field, recursive invocation, cross-project and write rejection | Exact families and bounded schemas; no host dispatch on rejected gateway operations. |
| LlmVbeTools.Privacy | Bound/unbound reads, exact grants and aliases, duplicate names, saved paths and malformed inventories, read versus edit permissions, independent and shared tools, filtered inventories and debug locations, live snapshot failure | No unapproved data leakage; failed aliases stay denied; unrelated native location fields are removed. |
| ChatWindow.ProjectAccess | Missing/busy state, legacy context alternatives, owned provider disposal, scope/inventory errors, non-array data, bound and ambiguous project filtering, cancel and accept | Cancel preserves the conversation; accept creates a new conversation with exactly the chosen grants; migration keeps local history while clearing provider context. |
| ProjectAccessWindow and Designer | Design/runtime creation, null and case-insensitive grants, separate shared permission, checked choices, owned thread dispatch, disposed appearance callback, Dispose true/false and missing container | Selection and shared permission are independent; modal buttons expose OK/Cancel; disposed controls are not updated. |
| ChatWindow.Budget | Gateway labels and malformed results, interrupted tools, pause state and summaries, busy/draft/session controls, every resume configuration mismatch, first-history alternatives, deduplicated pending responses, malformed native tool calls, cancellation boundaries, provider callbacks, missing text, applied changes and suspended verification, disposal during reply, oversized saved context | Completed calls are not replayed; a provider response is recorded as requiring state inspection; interruptions remain explicit; busy/client are released after faults; suspended verification finishes before the turn is closed. |

All provider traffic uses owned in-memory handlers or transport delegates. The bridge tests use private local named pipes and an injected native command boundary. No Excel or SolidWorks instance, user window, authenticated provider process, or real network request is used.

The only production change routes ProjectAccess through the existing ShowModal delegate; its default still calls Form.ShowDialog.

The oversized-history scenarios assert faulted tasks. A saved context above the serializer limit releases busy/client before persistence fails. A window closed during a provider reply skips disposed UI updates and rethrows the pending recovery exception; its stack must identify CompletePendingToolResponses. Neither case is recorded as a successful continuation.

## Measured result

The isolated XPlat measurement passed 111 tests, with no skipped tests. Production coverage has no exclusions; only the separate ProviderTests executable is excluded.

| Source | Lines | Branches |
| --- | --- | --- |
| BridgeRequestReader.cs | 35/35 | 18/18 |
| ChatWindow.Budget.cs | 144/144 | 132/132 |
| ChatWindow.ProjectAccess.cs | 46/46 | 38/38 |
| LlmVbeTools.Catalog.cs | 55/55 | 110/110 |
| LlmVbeTools.Privacy.cs | 98/98 | 92/92 |
| ProjectAccessWindow.cs | 23/23 | 10/10 |
| ProjectAccessWindow.Designer.cs | 72/72 | 4/4 |

Report: artifacts/coverage/pr10-access/9269526e-adaf-481c-8b21-a8685554fa8c/coverage.cobertura.xml.

BuildOutputRoot is an absolute path under this worktree's artifacts/build/pr10-access. Instrumentation runs are serialized. The test filter includes BridgeRequestBudgetTests, BridgeRequestReaderBoundaryTests, ToolCatalogTests, CatalogBoundaryTests, LlmProjectPrivacyTests, ProjectPrivacyBoundaryTests, ProjectAccess and ChatWindowStateTests. The collection setting is Exclude=[ProviderTests]*.

The mirror layout gate passes: 206 dedicated mirrors for 265 production files. This targeted measurement does not replace the global test gate or a real host acceptance run.