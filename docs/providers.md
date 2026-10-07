# AI providers

[Documentation](README.md)

Provider choice is independent of the application hosting the VBE and independent
of the GitHub account used for source versioning. VBAi connects to your provider;
it does not supply an account, an API credit balance or a model-usage entitlement.

## Configure a connection

Open VBAi settings, choose the provider, enter its credentials and endpoint where
required, and save. Select the model and any supported reasoning effort in the
chat. A blank credential field preserves a saved key; use the explicit removal
option to clear it. Saved settings take precedence over environment fallbacks.
Restart the host after changing environment variables inherited by its process.

Keys are stored with Windows DPAPI for the current user. Remote HTTP-provider
endpoints require HTTPS; HTTP is accepted only on loopback addresses. Embedded
URL credentials and fragments are rejected by both settings and transport
validation. Automatic HTTP redirects are disabled. Do not paste credentials into a conversation or issue.

## Connection identity

Each client captures an internal connection profile that binds the selected
provider, model, protocol, authentication mode and endpoint. Request headers use
that credential snapshot and may only target the captured origin: scheme, host
and port must match. This prevents a catalogue request or a later settings change
from reusing credentials for a different connection. DPAPI storage, environment
fallbacks and the existing CLI sign-in flows remain the credential sources.

The separation follows the provider/model approach reviewed in
[OpenClaw’s AI transport sources](https://github.com/openclaw/openclaw/tree/28c5d198851330906aa6feafc3effc3682021a15/packages/ai/src).
VBAi routes Codex through app-server, Copilot through its SDK protocol, compatible
HTTP providers through Chat Completions, Claude through Messages and Bedrock
through Converse. The profile records those implemented choices; it does not
convert account OAuth credentials into API keys or add a gateway.

HTTP model discovery derives `/models` only from a recognized `/chat/completions`
or `/messages` path. It retains the configured origin, any path prefix and the
query string. Claude catalogue pagination replaces only the `after_id` cursor
while keeping other query fields. Ollama uses its existing `/api/tags` endpoint
on the configured origin with the configured query. Connections with an explicit model list retain
that workflow; an unrecognized catalogue path reports an error.

## CLI-backed providers

### Codex

Codex is the default integration and uses `codex app-server` with CLI-managed
ChatGPT authentication. In this mode, an OpenAI API key is not required. Model
availability and reasoning options come from the provider catalog rather than a
hard-coded list.

VBAi sets the child process's `CODEX_HOME` to
`%LOCALAPPDATA%\VBAi\Providers\Codex`. Sign in from VBAi when necessary; being signed
in to another CLI data directory does not establish this connection. Existing
personal CLI authentication files are not copied implicitly.

Model discovery initializes the transport and account without opening or resuming
a conversation. A missing historical thread therefore does not block the model
catalogue. Sending a turn still requires that thread to resume successfully; a
failed resume does not silently create a replacement conversation.

Dynamic tool requests require the native turn identifier. The accepted identity
comes from the turn-start response; notifications arriving earlier wait for that
response and cannot establish their own authority. Tools and activities are
bound to their original turn generation and project, including callbacks queued
before Stop or completion. Terminal receipt closes tool admission immediately.
Late results cannot update a later turn or reopen a completed one.
VBAi records a local fingerprint of the developer instructions accepted for each
Codex thread. A resumed thread receives updated instructions only when that
fingerprint differs; unchanged turns do not send the full text again. Threads
created before this tracking was added receive the current instructions once on
their next resume. A failed update blocks the turn instead of continuing with
instructions whose version is uncertain.

Use `VBAi_CODEX_CLI` to specify the native `codex.exe` when automatic discovery
cannot locate it. Otherwise VBAi checks known installed locations and PATH.
A `.cmd` launcher is not accepted by this shell-free transport.

See the official [authentication](https://developers.openai.com/codex/auth/)
and [app-server](https://developers.openai.com/codex/app-server/) documentation for
provider-side behavior. Subscription access and API billing are distinct; no
fixed plan entitlement or quota is promised here.

### GitHub Copilot

Use the native Copilot CLI and its own sign-in flow. VBAi sets `COPILOT_HOME` to
`%LOCALAPPDATA%\VBAi\Providers\Copilot`; `VBAi_COPILOT_CLI` can override the path to
`copilot.exe`. This connection does not extract credentials from Visual Studio.

The implemented transport accepts protocol versions 2 and 3 and rejects other
versions with a diagnostic. Native shell/file/network/MCP permission requests are
refused in this transport; registered VBA tools still pass through VBAi's guards.
GitHub source-control authentication remains separate.

Permission and tool-result RPC errors settle their captured active turn before
a later idle event can publish success. Duplicate errors after an accepted RPC
result are ignored. Deferred error and activity callbacks cannot fault a
replacement turn; a failed activity receipt retains its terminal status through
client disposal. Activity publication errors preserve the primary RPC failure.

## HTTP and local providers

| Provider | Default endpoint or required endpoint | Credential fallback |
| --- | --- | --- |
| OpenAI API | `https://api.openai.com/v1/chat/completions` | `OPENAI_API_KEY` |
| Claude | `https://api.anthropic.com/v1/messages` | `ANTHROPIC_API_KEY` |
| Gemini | `https://generativelanguage.googleapis.com/v1beta/openai/chat/completions` | `GEMINI_API_KEY` |
| Mistral | `https://api.mistral.ai/v1/chat/completions` | `MISTRAL_API_KEY` |
| DeepSeek | `https://api.deepseek.com/chat/completions` | `DEEPSEEK_API_KEY` |
| OpenRouter | `https://openrouter.ai/api/v1/chat/completions` | `OPENROUTER_API_KEY` |
| Grok | `https://api.x.ai/v1/chat/completions` | `XAI_API_KEY` |
| Groq | `https://api.groq.com/openai/v1/chat/completions` | `GROQ_API_KEY` |
| Ollama | `http://localhost:11434/v1/chat/completions` | Optional key in settings. |
| LM Studio | `http://localhost:1234/v1/chat/completions` | Optional `LM_STUDIO_API_KEY`. |
| Custom OpenAI-compatible | Supply the full Chat Completions URL. | Optional `VBAi_CUSTOM_API_KEY`. |
| Azure OpenAI | `https://<resource>.openai.azure.com/openai/v1/chat/completions` | `AZURE_OPENAI_API_KEY` or `AZURE_OPENAI_ENTRA_TOKEN`. |
| Amazon Bedrock | `https://bedrock-runtime.<region>.amazonaws.com` | `AWS_BEARER_TOKEN_BEDROCK`. |

Endpoint fallbacks use `VBAi_<PROVIDER>_ENDPOINT`, with provider identifiers
`OPENAI`, `OLLAMA`, `CLAUDE`, `GEMINI`, `MISTRAL`, `DEEPSEEK`, `OPENROUTER`,
`LMSTUDIO`, `CUSTOM`, `AZURE`, `GROK`, `GROQ` and `BEDROCK`.

Custom, Azure and Bedrock connections take an explicit model list, one identifier
per line. The fallbacks are `VBAi_CUSTOM_MODEL`, `VBAi_AZURE_MODEL` and
`VBAi_BEDROCK_MODEL`. Azure uses deployment names. Its Entra mode accepts a supplied
bearer token; interactive Entra sign-in and automatic renewal are not implemented.
Replace an old API key when switching to token mode.

Bedrock uses the native Converse API and its bearer API key, not IAM access/secret
keys, SigV4 or an AWS profile. Supply a model or inference-profile identifier/ARN
available in the selected region. Custom connections currently have one profile.

## Model capabilities and transport limits

The model tooltip reports explicit catalogue declarations for tool calling,
reasoning and vision as supported, not supported or unknown. Missing or malformed
metadata remains unknown; model names do not determine these states. A declared
absence of tool calling uses text-only requests. Unknown support retains the
existing guarded tool workflow. The declared capabilities describe the model;
commands and request fields follow the implemented provider transport.

These declarations are immutable connection metadata. Manual model identifiers
retain unknown capabilities. They do not add persisted settings or establish live
qualification.

| Catalog | Explicit declarations read by VBAi |
| --- | --- |
| OpenRouter | `supported_parameters` lists tool and reasoning parameters; `architecture.input_modalities` lists image input. |
| Mistral | Native booleans `capabilities.function_calling` and `capabilities.vision`; reasoning remains unknown. |
| GitHub Copilot SDK | Native booleans `capabilities.supports.reasoningEffort` and `.vision`; tool support remains unknown. |
| Codex app-server | `supportedReasoningEfforts` and explicitly returned `inputModalities`; tool support remains unknown. |
| Other catalogs and manual lists | Unknown unless their supported declaration schema is implemented. |

An explicit complete list without a capability reports unsupported; an absent
list is unknown. Codex efforts consisting only of `none` do not declare reasoning
support. Reasoning metadata describes the provider's declared reasoning controls,
not a guarantee that a public reasoning block will be returned. Vision metadata
does not add an image-input workflow. VBAi does not enable new HTTP reasoning
parameters or derive effort levels from model names.

The declarations follow the current primary schemas for
[OpenRouter models](https://openrouter.ai/docs/guides/overview/models),
[Mistral models](https://docs.mistral.ai/api/endpoint/models),
[Copilot SDK model types](https://github.com/github/copilot-sdk/blob/main/nodejs/src/types.ts)
and [Codex model-list schema](https://github.com/openai/codex/blob/main/codex-rs/app-server-protocol/schema/json/v2/ModelListResponse.json).
A listed model may still fail to produce correct tool arguments or observe a
request. Test a harmless explanation and an inspection operation before enabling
agent changes. Local model quality and latency depend on the selected server/model.

Compatible providers and Claude stream text when supported; fragmented tool
arguments are assembled before execution. Truncated or incomplete responses must
not execute partial calls. Bedrock currently returns a complete Converse response,
not ConverseStream. Claude's implemented response limit is 8,192 tokens.

Synthetic UI qualification can retain bounded stream metadata: chunk counters,
whether text or tools were received, the final marker, a filtered terminal reason
and a complete/empty/error outcome. These metadata exclude prompts, response
content, tool names and arguments; they do not add a retry or change parsing.
An empty terminal response remains distinct from a truncated response or a
tool-only round. The selected Ollama configuration passes its stated live test
scope, including the six installed Office assistants. Historical intermittent
empty responses remain unresolved; see [qualification](release-qualification.md).

OpenAI requests include `store=false`; this is not a universal retention setting
for every provider or a guarantee of zero retention. Provider-side data handling
is governed by that provider and the selected account. See [privacy](privacy.md).

### Ollama configuration boundary

VBAi sends Ollama the selected model, messages, tool schemas and the
stream flag through `/v1/chat/completions`. Ollama's settings view also exposes
optional temperature and top-p overrides. Blank fields omit the corresponding
request fields and preserve the server's behavior. Temperature must be finite
and in `[0,2]`; top-p must be finite and in `(0,1]`. Entries accept an invariant
decimal point or the current decimal separator. The overrides persist as
`OllamaTemperature` and `OllamaTopP`; older settings leave them unset. A client
captures these values when constructed, and other providers ignore them.
Seed and context-size overrides are not sent. Model presence in the catalogue is not a guarantee of
correct tool arguments or adherence to a request to answer without tools.

On the tested diagnostic backend version `0.34.4`, the
[OpenAI request converter](https://github.com/ollama/ollama/blob/v0.34.4/openai/openai.go#L644)
sets omitted temperature and top-p to 1. Changing only those defaults in a
Modelfile does not change what this converter supplies. Ollama's
[context setting](https://docs.ollama.com/context-length) is separate; its
[OpenAI compatibility guide](https://docs.ollama.com/api/openai-compatibility#setting-the-local-context-size)
describes configuring a model's context rather than sending `num_ctx` as an
OpenAI chat field. Check the effective server settings and model capabilities
when diagnosing a local provider. The qualification profile uses a separately
selected model, an explicit server context and explicit request sampling; it
does not overwrite personal configuration. See [testing](../tests/README.md)
and [recorded validation](test-coverage.md) for the actual observed results.

The accepted isolated profile selects `qwen2.5:7b-instruct`, temperature `0`,
top-p `0.8`, server context `8192` and one parallel request. It passes exact
synthetic tool arguments, shown chat streaming/cancellation/recovery and real
read-only Excel inspection on the recorded candidate. These are explicit
qualification settings, not new defaults for existing personal profiles; the
candidate has not been deployed to the installed add-in. Other models, server
versions and embedded host paths require their own acceptance evidence.

## Validation

Transports are covered by simulated HTTP/CLI tests. The recorded live OpenRouter
check used synthetic text and a dummy tool, not a user's VBA project. Other
provider/account combinations are not implied to have been authenticated and
qualified merely because their transport exists. Consult
[recorded validation](test-coverage.md) and [testing](../tests/README.md).
