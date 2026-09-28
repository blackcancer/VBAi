# Native palette dialog matrix

Prepared before execution. VBE/CommandBars are managed objects; every HWND is an owned synthetic native #32770/tab/list/combo/button window on the disposable STA. UIA reads real list items; message notifications update an in-memory palette. No Office command, profile settings, SendKeys or user-window action.

| Area | Cases |
| --- | --- |
| Owner/command | absent/hidden/disabled owner; null/disabled command; command exception |
| Discovery | owned new dialog, existing dialog excluded, ambiguous candidates, no candidate/time out; exact native class/buttons/tab requirement |
| Page guards | disabled tabs, button-style tabs, zero/oversized count, cancellation, missing page, selected-page mismatch, duplicate/missing selectors |
| Read/write | ten real UIA list names, category count rejection, each selector count, invalid category select, invalid color select, retained selection rejection, successful notifications |
| Visit | read-only cancel, successful update accept and readback, invalid desired rows, category rename, callback exception, controls mismatch |
| Close/wait | missing button, post failure, delayed close, never closes, canceled open, join not confirmed |
| Ownership | visible same-process owner chain accepted; unrelated/preexisting/hidden windows rejected |

Timeout policy and post/join seams retain production defaults and are restored after each fixture. Shorter timeouts exercise real bounded loops; post/join failures model environmental outcomes while preserving exception and cleanup checks.

Qualification: all production lines and helper branches were executed. Visit retains one defensive branch at its final dialog != IntPtr.Zero check: after confirmed worker completion without failure, discovery has already established a nonzero dialog. The zero-success branch is not fabricated by corrupting the closure. The guard is preserved at the root agent's explicit request. Native default timeout/post/join behavior remains unchanged.
