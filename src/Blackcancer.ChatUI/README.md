# Blackcancer.ChatUI

Shared source module for VBAi and SolidAi, targeting .NET Framework 4.8 / AnyCPU.
The source directory is vendored unchanged in both repositories and both hosts
bind their workspace and Markdown rendering to it. The host owns
chat sessions, tools, native APIs, culture selection and theme selection.
This module has no COM, VBE, SOLIDWORKS, VBAi or SolidAi.Core dependency.
Its local `.gitattributes` preserves exact source bytes across Git checkouts;
host-specific line-ending rules must not invalidate the parity manifest.

`Presentation.ChatMarkdownEngine.Parse` returns immutable blocks and spans, with
original Markdown, fence-free code and structured table cells. It uses the same
Markdig 0.37.0 pipeline for both hosts. HTML is inert text.

`Views.ChatMarkdownView` owns Designer templates for prose, code and tables.
Its child block instances are created dynamically within that fixed template.
All public controls retain `.cs`, `.Designer.cs` and `.resx` files.

```csharp
var transcript = new Blackcancer.ChatUI.Views.ChatMarkdownView();
transcript.Services = new Blackcancer.ChatUI.ChatServices {
    Translate = hostTranslate,
    CopyText = hostCopy,
    OpenLink = hostOpenLink
};
transcript.Style = hostChatStyle;
transcript.ShowMarkdown(markdown, hostError, tokenActions);
transcript.AppendDelta(delta);          // Call after the host's stream batch.
transcript.CompleteMarkdown(finalText); // Reconciles final block structure.
```

These calls belong on the control's UI thread. The host batches provider events
(typically 50 ms), manages cancellation and flushes completion before releasing
the live message. There is no global host service or host DTO in this module.
`Runtime.TranscriptUpdateBatch<TKey>` supplies a host-independent pending-work
helper; scheduling and the UI timer belong to the host.

`UpdateMarkdown` skips identical source and reuses exact blocks across structure
changes. It updates changed blocks of the same kind in place. Selection, native
text scroll, table scroll and code disclosure choices survive these updates.
RichEdit `EM_REQUESTRESIZE` supplies document height, independent of caret scroll;
controls bound document height and provide scroll bars for overflow.

Code headers retain a language caption and Copy button while collapsed. Blocks
with more than twelve lines initially collapse. User choices persist as blocks
stream and as final rendering inserts or removes neighboring content.
Tables use a read-only virtual `DataGridView`, bounded column-width sampling,
horizontal scrolling, alignment metadata and quoted TSV copying. Full cell data
is retained even when the visible cell or width sample is bounded.

`CopySource` optionally overrides copying of the complete original Markdown;
`SourceText` returns that source verbatim. The default services use the native
clipboard and permit absolute HTTPS links only. Each host injects its own
translation, copying, navigation and style callbacks per instance. Fonts in a
`SharedChatStyle` remain owned by its host.

The prose measurement and width retention originate from SolidAi's corrected
transcript controls. The caption-only disclosure update originates from VBAi's
working tree. Shared Markdown traversal is adapted from the MPL-2.0 VBAi/SolidAi
native renderers; retain license/provenance files when vendoring.

Behavioral coverage is in the host test projects. Source-level integration and
managed tests do not establish acceptance inside a native SOLIDWORKS or VBE host.

The source-vendoring tool is `tools/Sync-SharedChatUI.ps1` in the integration
repository. After changes and Designer round-tripping are stable, generate the
baseline manifest, then verify or explicitly synchronize the intended peer:

```powershell
& .\tools\Sync-SharedChatUI.ps1 -Mode Manifest
& .\tools\Sync-SharedChatUI.ps1 -Mode Verify -PeerRepository 'E:\Development\PeerRepository'
& .\tools\Sync-SharedChatUI.ps1 -Mode Sync -PeerRepository 'E:\Development\PeerRepository'
```

`Manifest` writes `build/SharedChatUI.sources.json`, recording the exact module
file set and SHA-256 bytes. All module project files, sources, Designer files,
resources, documentation and licenses participate; `bin` and `obj` directories
and local IDE/configuration files (`*.user`, `*.suo`, `*.local.*`, `.env*`)
are excluded. `Verify` requires matching source files and matching manifests in
both repositories. This check detects source forks even if the UI looks similar.

`Sync` requires an explicit existing peer directory. It checks the source against
its manifest and the peer against its own previous manifest before updating files,
then copies the exact module bytes and source manifest. An absent peer module can
be initialized. An existing module without a manifest can only be registered if
its sources already exactly match the source baseline. Divergent local peer
changes are refused. Source removals or case-only renames requiring peer cleanup
are refused; the tool never deletes files. Paths must stay inside the specified
repositories, and reparse points are rejected. Neither Git nor a build runs.
If files change concurrently after preflight, synchronization stops and leaves
any completed writes visible for review, without retrying or cleaning them up.
