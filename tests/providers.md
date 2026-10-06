# Provider qualification

[Testing overview](README.md) · [Provider configuration](../docs/providers.md)

## Q028 complete Office assistant campaign

`tools/tests/Invoke-Q028Qualification.ps1 -Prepare -EvidenceRoot <fresh-absolute-root>
-ModelRoot <existing-absolute-model-store> -BuildOutputRoot <isolated-build-root>`
freezes the complete provider and six-Office-host matrix before any request or
native launch. The selected profile uses the already installed
`qwen2.5:7b-instruct` model, CPU inference, context 8192, one parallel request,
temperature 0 and top-p 0.8. Model manifest/blob and product/test/executable
hashes are pinned. Preparation never downloads or substitutes a model.

The matrix contains managed prerequisites, the strict synthetic tool roundtrip,
cancellation/recovery, detached streaming UI, then the embedded assistant in
Excel, Word, PowerPoint, Access, Publisher and classic Outlook. Every native bank
requires streaming while busy, one Stop, visible cancellation, a completed next
send, an unprompted native marker through exactly one `read_module`, unchanged
source/references and normal owned-host exit. No VBA runs or mail is sent.
A failed or uncertain bank stops later banks; never retry an unsettled action.

The original private-desktop mode retains its guard. For an explicitly approved
real-desktop campaign, prepare with `-MainDesktopAuthorized` and dispatch using
`tools/tests/Invoke-Q028Main.ps1 -EvidenceRoot <prepared-root>
-ExpectedPlanSha256 <reviewed-hash>`. Its same-user Limited x64 STA worker runs on
`WinSta0\Default`; both private-desktop descriptor variables must be empty.
The worker sets `VBAi_RUN_OLLAMA_OFFICE_MAIN_DESKTOP_TESTS=1` only for native banks;
Word also uses its established main-desktop bootstrap and the frozen product
MVID/SHA pins required by that bootstrap. No desktop switch occurs.
Pre-existing hosts and personal Outlook VBA projects refuse qualification.

Discovery admits the exact embedded ActiveX site once. During an unsettled send,
the guard verifies the original HWNDs, PID/thread, native classes, visibility
and parent/owner chain. It never rediscovers a replacement container or rereads
an unavailable parent UIA provider. Combo choices and native button delivery
remain single actions requiring exact control identities and readback.

After terminal UI proof, the original observer drops its managed UIA references
before native shutdown. The selected Word bank also uses the owned, settled
testhost collection diagnostic. This does not prove native RCW release; the
original Word process handle must still observe normal exit within its unchanged
five-second bound. Each other host retains its own frozen shutdown deadline.
An exit after the applicable deadline remains a failed bank.

Before any model request, the original isolated Ollama backend enters a private
Windows job. Its future calculation workers inherit that job. Once requests
settle and Office exits, shutdown terminates only that synthetic job and verifies
empty kernel membership, alongside the original backend exit handle. Stopping
only `ollama.exe` is insufficient: orphan `llama-server.exe` workers can retain
large committed buffers. There is no kill-on-close flag, process-name sweep or
shutdown of a personal/default backend; uncertain native work is retained.

`tools/tests/Review-Q028Wire.py --self-test` validates refusal oracles offline.
Its campaign review binds exact wire arguments/results, the original worker/host
lifecycle and settings/registration restoration. Later passes do not explain
earlier empty responses. Results remain candidate/profile-specific and belong
in [recorded validation](../docs/test-coverage.md).

## Synthetic Ollama scenarios

`VBAi_RUN_OLLAMA_TESTS=1` enables `TestCategory=Ollama` against the loopback
server through the production HTTP client. `VBAi_TEST_OLLAMA_MODEL` selects an
already-installed model (default `qwen2.5:7b-instruct`). The shared selector applies
to the headless, detached UI and native Excel scenarios. An absent override uses
the default; outer whitespace is trimmed, while an empty explicit override,
internal whitespace or control characters are refused. It never loads or rewrites
personal provider configuration, downloads a model or silently substitutes another
model. These synthetic scenarios cover
streamed text, a harmless tool roundtrip, cancellation and a subsequent request;
they do not execute native VBE tools or read saved provider settings. Remove the
opt-in variables after the run.

The headless HTTP, detached chat UI and native Excel cases accept
`VBAi_TEST_OLLAMA_ENDPOINT` when an
owned local server uses another port. Its default remains
`http://127.0.0.1:11434/v1/chat/completions`. An override must be a canonical
`http://127.0.0.1:<port>/v1/chat/completions` URL without credentials, query or
fragment. Optional synthetic wire capture is restricted to that exact server's
chat and `/api/tags` routes and records the selected port. Retain the backend
version and model digest; another port does not prove the default port is usable.

The qualification backend profile is an already-installed `qwen2.5:7b-instruct`
model with `OLLAMA_CONTEXT_LENGTH=8192`, `OLLAMA_NUM_PARALLEL=1`,
`OLLAMA_NO_CLOUD=1` and `OLLAMA_NOPRUNE=1`, passed only to an owned local backend.
Attest the selected manifest/blob hashes and the backend's cloud-disabled startup
state before running the enabled scenarios. Do not change personal/global
environment. The shared test profile explicitly requests temperature `0` and
`top_p=0.8` through isolated `LlmSettings`, selected by
`VBAi_TEST_OLLAMA_TEMPERATURE` and `VBAi_TEST_OLLAMA_TOP_P` overrides when needed.
The overrides use finite invariant-culture numbers: temperature is in `[0,2]`,
and top-p is in `(0,1]`. The scenarios report the actual model, endpoint and
sampling before requests; they do not rewrite user settings or change prompts or
assertions. Production Ollama sampling properties remain nullable by default;
old settings therefore retain the backend's previous request behavior, and other
providers ignore these Ollama-only fields. This preparation is distinct from a passing
provider qualification: the original scalar, tool, UI and native readback
assertions remain required. The qualification requests do not add `logprobs` or
`top_logprobs`; controls using those diagnostic fields remain separate protocol
experiments and cannot replace the unmodified qualification request.

Passive synthetic body capture requires a separate explicit flag. For the two
headless scenarios, set `VBAi_OLLAMA_HEADLESS_CAPTURE_WIRE=1` and optionally
`VBAi_OLLAMA_HEADLESS_RESULTS` to an absolute local evidence directory (default:
`ollama-headless-diagnostics` under the test output). For the detached UI scenario,
the existing `VBAi_OLLAMA_UI_CAPTURE_WIRE=1` and `VBAi_OLLAMA_UI_RESULTS` variables
retain their behavior and `ollama-ui-diagnostics` default. The scenario opt-in is
still required; enabling capture alone does not run a model.

The shared `OllamaSyntheticWireCapture` helper requires the real production
`HttpClientHandler` with redirects disabled and the exact selected IPv4 loopback
port, POST chat or GET catalogue route. It observes buffered synthetic requests
and response bytes only as the production reader consumes them. Each body capture
is bounded to 1 MiB; read summaries distinguish observed EOF, early disposal,
in-flight reads, read errors and truncation. SSE `[DONE]` can stop the production
reader before physical EOF, so a false EOF flag alone is not a truncated-response
claim. Headers, credentials and personal history are not recorded. Diagnostic
write/serialization/wrapping failures do not replace the HTTP outcome.

The headless tool scenario records raw argument JSON, parsed argument type and
the marker's type/value before its existing scalar assertion. A nested object is
preserved as an object; it is never flattened or accepted as the expected string.
Evidence includes request intents, completion/error phases and the actually loaded
product MVID. No native tool is dispatched, and later evidence does not establish
the cause of an earlier response without its own captured wire.

For a detached helper-only regression batch, use the filter
`FullyQualifiedName~OllamaSyntheticWireCaptureTests|FullyQualifiedName~OllamaQualificationEndpointTests|FullyQualifiedName~OllamaQualificationModelTests|FullyQualifiedName~OllamaQualificationProfileTests`.
These tests use memory streams and explicit fake production-client responses;
they do not open sockets, run models or launch a native host. Remove the capture
and scenario opt-ins after real qualification.

`VBAi_RUN_OLLAMA_UI_TESTS=1` enables `TestCategory=OllamaUi` with the same
model selector. It shows the real chat controls and checks send, rendered
streaming, Stop and a subsequent completed response through the production
loopback HTTP client. Settings and history are isolated; the VBE project is
simulated and native tools are refused. This qualifies a detached chat workflow,
not Office or SOLIDWORKS integration. Run it separately from native host UI tests
to avoid competing for focus, then remove its opt-in variable.

`VBAi_RUN_OLLAMA_EXCEL_TESTS=1` enables `TestCategory=OllamaExcel` using the
same guarded loopback endpoint and already-installed model selectors. It creates a disposable Excel module containing a
random marker absent from the prompt, dispatches the model's `read_module` call
through the real project-bound VBE tools, and verifies the final answer and
unchanged source and normal exit of its owned Excel process. No macro runs.
This covers production provider/tools/session code dispatching through native
Excel COM from the test process, not the installed bridge or embedded assistant UI.
Run it separately from other native host tests
and remove the opt-in afterwards.

## Existing-account and exported-source scenarios

Existing-account checks require both `VBAi_CONNECTED_PROVIDER_TESTS=1` and
`VBAi_CONNECTED_SOURCE_TESTS=1`. The Git read check additionally requires
`VBAi_TEST_GITHUB_MANIFEST` identifying an explicitly authorized synthetic private
repository (URL, ID, verified ownership/private state and expected main commit).
The Codex check requires `VBAi_TEST_SOLIDWORKS_MANIFEST` with `OwnedDisposable`,
absolute `Path`, `FileSha256`, `Module`, `ModuleSha256`, `Marker`, `Pid` and `Mvid`.
It accepts only the manifested disposable SWP under the qualification artifacts,
uses an already-connected account and never signs in or copies authentication state.

Native UserForm GitHub qualification is a separate explicit scope:
`VBAi_RUN_USERFORM_GITHUB_TESTS=1`, `VBAi_RUN_EXCEL_TESTS=1`, the Git manifest and
an absolute `VBAi_TEST_USERFORM_GIT_OUTPUT`. The fixture is constrained to the
maintainer-authorized retained qualification repository and creates a dedicated
branch containing synthetic exports. It verifies controls, code and FRX bytes,
retains backups and captures the owned designer windows for visual review.
`VBAi_RUN_USERFORM_CORRUPTION_TESTS=1` separately enables the local-only malformed
FRX diagnostic; it must not publish corrupt content. Run either scenario only
while it owns the desktop. Normal host shutdown is part of acceptance; a verified
transfer does not excuse a subsequent crash. These fixtures are not enabled by
ordinary connected-account opt-ins.


## Reporting provider acceptance

Keep protocol simulations, real-provider traffic, detached chat and installed
embedded assistants distinct. Record model/server/profile, exact synthetic tool
arguments/results, cancellation, subsequent response and resource cleanup.
Production prompts, credentials and personal histories are excluded from
publication. The selected Q028 result is in
[recorded validation](../docs/test-coverage.md#q028-ollama-and-embedded-office-assistant-qualification-2026-10-06).
