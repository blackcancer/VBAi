# User manual

[Documentation](../README.md)

The offline user manual is authored in French first. It explains the product's
workflows for users who already know VBA, AI assistants and GitHub. The worked
example uses a disposable Excel project, real Codex/Luna requests and a separate
request to add and run tests before the Git workflow. Other help languages follow review
of this first edition; the language of maintained repository Markdown stays English.

## Read and build

The source is [the French manual](fr-FR/manual.json). It produces ordinary HTML
for review and an optional Windows HTML Help archive with contents, keyword index
and full-text search. Python 3.10 or later is sufficient to produce HTML:

```powershell
python tools/docs/build_help.py
```

For CHM, supply the path to a local Microsoft HTML Help compiler:

```powershell
python tools/docs/build_help.py --compiler 'C:/path/to/HTML Help Workshop/hhc.exe'
```

The builder does not download or install a compiler. Generated output goes to
`artifacts/help/fr-FR/`, including `VBAi.fr-FR.chm` when compilation is requested.
Inspect compiler logs and the actual archive before distributing it. An old
archive is removed before a new compilation so it cannot mask a failed build.

When the archive exists before building VBAi, the project copies it to
`Help/VBAi.fr-FR.chm` beside `VBAi.dll`. **About VBAi → User guide** opens the
local entry page. Ordinary form help requests route to the relevant chapter.
An absent manual produces an explanation, without an automatic web redirect.
This source-build packaging does not establish an installer or published release.

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
