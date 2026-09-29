# Renderer lifecycle matrix

Prepared before execution. Fixture holds scoped static fields and managed callbacks; a synthetic nonzero module prevents loading. No native DLL extraction or hook installation.

| Method | Complete branch cases | Assertion |
| --- | --- | --- |
| Start | success, nonzero error, callback exception; cleanup success/failure | start argument and count; active transition; explicit stop count and preserved module |
| Register | inactive, success, unsupported 50, failure | no inactive calls; HWND preserved; only error requests stop |
| Refresh | inactive, success, failure | no inactive calls; only error requests stop |
| Describe | missing query, query failure, populated status | exact diagnostic and ABI structure size |
| Hash | known SHA256 vector | payload hash compatible with production cache key |

EnsureLoaded unload/extract/export paths remain separate unqualified work. This lot covers lifecycle after a validated loaded module; it cannot prove native hooks or real VBE behavior.
