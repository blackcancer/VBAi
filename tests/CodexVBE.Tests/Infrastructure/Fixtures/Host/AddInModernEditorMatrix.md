# AddIn modern editor orchestration matrix

Prepared before tests. Hosts, panes and tool windows are owned synthetic objects; editor windows are constructed normally with browser initialization held pending by the factory fixture. Drafts, settings and crash reports stay in owned temporary roots. No Excel, macros, WebView profile, native palette/settings or COM registration.

| Area | Cases |
| --- | --- |
| Active module | follow=false/true; missing active window; non-code window; missing code pane; project non-design mode; valid module identity |
| Modern editor | show=false absent/live/disposed; create/recreate; disposed native site reset; visible reactivate; hidden show with owned owner; already docked attach and focus |
| Navigation/menu | open real in-memory module; open refusal; show editor without module/with unavailable native module/host failure; /editor callback |
| Dock | first create/reuse; missing/disposed control; native creation refusal/invalid control/focus refusal; attach dimensions/focus; undock/re-dock; DockRequested callback |
| Cleanup | docked live/missing/disposed editor or control; native close refusal; repeated shutdown; resources disposed |
| Startup | settings failure; native dark enabled log with injected palette/theme; native owner/theme failure |
