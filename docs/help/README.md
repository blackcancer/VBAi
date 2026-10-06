# User manual

[Documentation](../README.md)

The offline manual explains VBAi workflows to users who already know VBA, AI
assistants and GitHub. Its worked example uses a disposable Excel project, real
Codex/Luna requests and a separate request to create and run tests before publishing
sources. French is the reviewed source edition. Translations are manually authored
prose; no translation service is used. Maintained repository Markdown remains English.

## Languages and source layout

The [French manual](fr-FR/manual.json) owns chapter IDs, interface assignments,
example code and illustration geometry. Each other language has a complete
`translations.json` in its culture directory. These catalogs translate all prose,
index terms, tables, captions, annotations and navigation labels in source order.
They carry the exact French source SHA-256; changing the reviewed source requires
reviewing the affected translations before building again. Missing or empty text,
missing topics and stale catalogs are build errors, with no silent French fallback
inside a translated chapter.

The guides cover the application's UI languages: English (`en-US`), French (`fr-FR`),
Spanish (`es-ES`), German (`de-DE`), Brazilian Portuguese (`pt-BR`), Italian (`it-IT`),
Japanese (`ja-JP`), Korean (`ko-KR`), Simplified Chinese (`zh-CN`), Traditional Chinese
(`zh-TW`), Russian (`ru-RU`), Arabic (`ar-SA`) and Hindi (`hi-IN`). Arabic pages use
right-to-left text while retaining original image coordinates and code direction.
The translations reuse authentic French screenshots and explicitly identify their
language; the application pixels are not translated or reconstructed.

## Build and package

Python 3.10 or later can generate ordinary HTML for review:

```powershell
python tools/docs/build_help.py --all
```

For compiled help, supply a local Microsoft HTML Help compiler:

```powershell
python tools/docs/build_help.py --all --compiler 'C:/path/to/HTML Help Workshop/hhc.exe'
```

A single edition can be built with `--source docs/help/en-US`. Review HTML and
compiler logs go to
`artifacts/help-staging/<culture>/`. After all requested compilations succeed, only
the compiled `VBAi.<culture>.chm` archives are published to `dist/help/`.
`--output` selects an alternative staging root for `--all`, or one staging directory
for a single edition. `--package-dir` selects the clean compiled distribution folder.
After reviewing a compiled batch and retaining its validation receipt, the
corresponding `artifacts/help-staging/` and temporary archive extraction folders
can be deleted. They are reproducible development output, not installation inputs.
Do not delete a frozen native qualification candidate merely because a later build
exists. The builder never downloads or installs a compiler. Inspect
compiler logs and the actual archives before distribution; an old archive is removed
before compiling its replacement.

Build the guides before building the solution. Both the add-in and updater copy all
compiled editions from `dist/help/` into their own `Help/` folder.
The resulting installation layout is:

```text
VBAi.dll
Help/
  VBAi.en-US.chm
  VBAi.fr-FR.chm
  ...
```

The updater carries the same folder beside its executable. `HelpOutputRoot` can
point MSBuild to a shared generated root when building from an isolated worktree:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build" -p:HelpOutputRoot="C:/distribution/help/"
```

**VBE Help / ? → VBAi · User guide** and **About VBAi → User guide** open the local
entry page. Form help requests route to their task chapter. The add-in selects the
VBE interface language; the updater uses its own UI culture. Regional variants use
the corresponding supported language. If its archive is absent, VBAi tries English,
then French. If no archive exists, it explains where to place the local files.
This packaging does not establish an installer or published release.

## Screenshot provenance

The [screenshot manifest](fr-FR/screenshots/manifest.json) records the real window
identity, capture time (or the original capture file timestamp), candidate assembly
SHA-256, dimensions and PNG hashes.
Use the application in its actual host. Do not wrap isolated UserControls in
invented windows, inject canned conversations or fabricate test results.

The help capture scenario records one existing window. It does not start a host,
send prompts, grant access, or operate a macro. Supply the exact host PID, HWND
and loaded candidate; choose a fresh output directory or capture ID:

```powershell
powershell.exe -NoProfile -STA -File tools/tests/Render-Ui.ps1 -Scenario Help `
  -AssemblyPath 'artifacts/build/VBAi/Debug/net48/VBAi.dll' `
  -HostProcessId 12345 -WindowHandle 67890 -CaptureId 'chat-model-selection' `
  -OutputDirectory 'artifacts/help-captures'
```

Replace the example PID and HWND with verified live identities. Native rendering
is the default. For a layered popup that cannot render through PrintWindow,
`-CaptureRenderer Screen` captures its actual visible rectangle; ensure no other
window covers it. Neither mode activates or moves the selected window. The helper
rejects a uniform image, but every capture still needs visual review for clipping,
legibility, unrelated content and correspondence with the documented operation.

Copy reviewed PNGs into the source tree and merge their records into the manifest.
A record must carry `provenance: live-interface` or `live-workflow`, its real
`windowTitle`, and `synthetic: false`. Preserve the original pixels. Section-level
`figure` records place illustrations beside their explanation; chapter galleries
are rejected. Choose `side`, `compact` or `wide` according to the illustration,
not a common full-page width. Optional numbered `callouts` use percentage rectangles
and explanatory text; HTML overlays leave the original screenshot untouched.
An optional `crop: [x, y, width, height]` focuses the figure on a relevant region
in source pixels; clicking the figure opens the complete original capture.
Crops must remain inside the source image and are never enlarged past their native size.

A provider response, source mutation, compilation and test execution are distinct
observations. Check generated components and terminal test reports before using a
capture as a successful example. Record a refused or incomplete action accurately.

## Maintain explanations and hints

Each fixed interface has one authoritative chapter in the manual's interface map.
Nested views are explained within their user workflow. Preserve the distinctions
between saving a host document, committing source locally, publishing remotely
and importing code. A screenshot must never imply a success not observed in the
real scenario it describes.

WinForms usage hints are in `src/VBAi/Ui/UiHelpHints.json`, with English and French
resource text in the existing localization catalogs. Owner-qualified entries
take precedence over shared names. Existing contextual hints remain authoritative;
Designer construction stays free of help viewers and live services. The WPF chat
composer also exposes its usage text as a tooltip and accessibility help.

Run the documentation checks and help-builder tests after changing these sources:

```powershell
python -m unittest discover -s tests/tools -p 'test_*.py'
python -m unittest discover -s tools/docs -p 'test_*.py'
python tools/docs/check_docs.py
```
