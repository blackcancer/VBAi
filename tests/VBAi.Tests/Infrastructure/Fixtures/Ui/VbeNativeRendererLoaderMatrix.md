# Renderer loader matrix

Prepared before execution. Cache IO is confined to a disposable temporary directory. Real embedded DLL loading is limited to export/ABI queries while inactive; Start is never called, so no hooks or Office process are involved. Failure callbacks are scoped and restored.

| Case | Required evidence |
| --- | --- |
| Valid embedded load | hash-named cache bytes exactly match embedded payload; ABI query inactive; delegates/module retained; no Start; module released by fixture |
| Already loaded | no resource/cache/load callback repeated |
| Unsupported host and missing resource | reject before cache mutation |
| Existing matching cache | no replacement; verified hash and successful load |
| Existing corrupt cache | reject before LoadLibrary |
| Move race winner | another valid destination exists; preserve winner; delete temporary file |
| Move failure without winner | propagate IOException and delete temporary file |
| Load failure | Win32Exception; no module retained |
| Missing export | EntryPointNotFoundException; module released; no partially assigned exports |
| ABI status failure and ABI mismatch | InvalidDataException; module released; inactive |
| Resolve null export and valid export | actual resolver result/rejection |

Native loading defaults remain unchanged. The test seams isolate host/resource/cache/native call boundaries; they do not exclude loader code or manufacture coverage by invoking retained native hooks.
