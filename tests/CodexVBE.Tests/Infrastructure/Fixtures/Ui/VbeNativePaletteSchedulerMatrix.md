# Palette scheduler matrix

Prepared before execution. Each service has a temporary recovery path, owned synthetic HWND, managed transaction/error callbacks, and a scoped static update lock/log snapshot. No Options command, settings file in user profile, uncontrolled message pumping. A final default-presenter case displays a synthetic MessageBox and dismisses only its uniquely marked dialog on the disposable owning STA.

| Operation | Cases | Evidence |
| --- | --- | --- |
| Constructor | invalid null/empty/filename version, valid temporary/default computed path | reject unsafe recovery names; default only computes path, performs no settings IO |
| Request | disposed, restore missing state, recovery present, first enable, same applied state, peer transaction active | timer enabled/disabled and requested value |
| ApplyPending guards | disposed, hidden, disabled, peer transaction active | zero transaction/error callbacks; queued timer retained when blocked |
| ApplyPending commit | enable/restore, target changed during callback, disposed during callback, nested peer invocation | callback arguments/count, applied value, update lock released, timer restarted only when needed |
| Failure | transaction exception | exact error callback and diagnostic, applied unchanged, lock released; translated default modal diagnostic closed by thread-scoped WM_COMMAND |
| Dispose | active timer/repeated disposal/post-disposal request | stops timer permanently and never opens Options |

Default delegates continue to use the production Change method and translated MessageBox. Only their environmental boundaries are substituted in the isolated fixture.
