# VbeBridgeClient transport matrix

- Existing process/command overloads preserve the owned process pipe name and 120-second response timeout.
- Server available: one JSON line, one object response, request counter exactly 1.
- Server created after at least one bounded connect timeout: retries connect only; one emission on success.
- No server after all connect attempts: null, no emission possible.
- Disconnect after receiving mutation: IOException, counter exactly 1, no reconnect/re-emission.
- Hold response after receiving mutation: configured TimeoutException, counter exactly 1, owned I/O disposed.
- Malformed JSON and valid JSON scalar response: parser/shape error, counter exactly 1, no retry.
- Invalid limits/name rejected before connection.

All pipes have random unique fixture-owned names. No Excel, VBE or production DLL is involved.