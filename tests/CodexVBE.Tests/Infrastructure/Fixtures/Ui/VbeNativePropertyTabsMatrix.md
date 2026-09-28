# Native PropertyTabs coverage matrix

Prepared before tests. Every native operation defaults to its original P/Invoke. Fixtures own real hidden tabs and memory DCs, restore all callback fields, and never drive host applications.

| Area | Cases |
| --- | --- |
| CanRender | zero/invalid/foreign thread/wrong class; unsupported style/extended style; count/image list/multi-row/child; each geometry API failure; zero/negative/oversize dimensions; origin and bounds mismatches; accepted layout |
| Creation and read | rejection/null renderer; native item failure/rectangle failure/empty width/height/null text/oversize text; fallback stock font; focus absent/present/hidden; enabled/highlighted |
| Message handling | unsupported/disposed/thread refusal; snapshot refusal; erase; paint DC present/absent and EndPaint; null print DC; visible and hidden check; print empty/nonclient/background/client; invalid DC saves |
| Drawing | caller state restored, background exact pixels; selected/unselected/hot/disabled/highlighted; focus each item; clipping/font/text operations |
| After native | hover each hit-test boundary and missing rectangle; track success/failure/repeat/leave/no-op; every state message; unsupported layout prevents repaint; owning thread refusal |
| Dispose | no tracking/tracking/cross-thread/invalid HWND/repeat; native tracking canceled only for owned live control |
