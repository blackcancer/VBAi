# Theme orchestration matrix

Prepared before execution. Scoped fixtures restore every static state/collection/delegate and process environment variable. Palette timers use disposable recovery paths without pumping or invoking Options. Preferred-mode and dark-window callbacks are managed fakes; no process-wide native mode is changed.

| Operation | Complete cases | Required state/resource evidence |
| --- | --- | --- |
| Initialize | zero owner; null VBE; panes missing/immediate/non-immediate; enumeration error; old palette; explicit experiment on/off; persisted enabled on/off | owner/caption discovery, old timer disposal, palette creation gate, requested state |
| Enable | no owner; first apply; already themed; apply failure after partial mutation; palette present/absent | exact apply count; cleanup after failure; original exception retained; palette request |
| Disable | native stop success/failure; palette present/absent | requests only after confirmed stop; owner/resource preservation on failure |
| Reset | stop refusal; event hook absent/present; subclasses/pending/brush; ordinary/dialog theme; preference unchanged/changed with present/absent delegates; flush present/absent; owner absent/present | collection clearing, hook/subclass removal, actual GDI brush deletion, previous preferred mode callback, owner retained |
| Disconnect | stop refusal/success; palette present/absent; discovered pane | resources preserved on refusal; successful disposal and owner/pane release |
| Apply guard | zero owner and foreign/missing thread owner | rejection before private API/process preference access |

Per-window native restore calls target only disposable fixture HWNDs owned by this test. Apply routing seam supplies only the application boundary so orchestration is exercised without process preference changes; native Apply remains separately measurable.
