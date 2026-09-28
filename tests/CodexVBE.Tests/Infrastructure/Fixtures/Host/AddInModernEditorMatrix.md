# AddIn modern editor PR10 workspace matrix

Prepared before batch. The HostUiScope owns a WinForms IsMdiContainer window and its real MdiClient HWND. Every editor is a native child of that owned document workspace. Browser initialization is held by its existing guard; draft/settings/crash storage is temporary. No user VBE or Office instance.

| Area | Cases |
| --- | --- |
| Active module | follow false/true; missing/noncode active window; absent pane; running/design mode; exact component |
| Editor workspace | absent/live/disposed getter; create/reuse/recreate; native GetParent matches owned MdiClient; show after hidden; resources and timer replaced; independent assistant site |
| Workspace visibility | active native form/designer versus code; client dimensions; explicit show; user close canceled while hosted |
| Creation refusal | owned host has no MdiClient; disposed rejected editor; cleared fields; logged notice; restoration with a new MdiClient succeeds |
| Navigation/menu | memory module open; native module refusal; missing host; null module; editor menu callback |
| Shutdown/startup | absent/live/disposed/already released workspace; repeated shutdown; unchanged settings/theme/bridge/logger error contracts |
| PR10 action routing | editor versus assistant fallback; managed/native attachment; missing readiness/current; script failure |