# User manual

[Documentation](../README.md)

The offline user manual is authored in French first. It explains the product's
interfaces through a fictional budget project, with task walkthroughs, expected
outcomes, comparisons and recovery guidance. Other help languages follow review
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

The [screenshot manifest](fr-FR/screenshots/manifest.json) records capture time,
candidate assembly SHA-256, dimensions and individual PNG hashes. These are
actual VBAi controls populated with fictional examples. They are not generated
illustrations and do not establish provider connectivity or native-host qualification.

To refresh the complete bank on Windows, explicitly choose the candidate and
output directory:

```powershell
powershell.exe -NoProfile -STA -File tools/tests/Render-Ui.ps1 -Scenario Help `
  -Culture fr-FR -AssemblyPath 'artifacts/build/VBAi/Debug/net48/VBAi.dll' `
  -OutputDirectory 'artifacts/help-captures'
```

The helper displays the controls outside the visible desktop without activation.
It renders RichEdit, hosted WPF and the real Monaco WebView2 surface with their
respective rendering APIs. Account state, messages and results are fixtures;
no provider request or Office/SOLIDWORKS macro is executed. Review every selected
image for legibility, populated content and clipping. A nonuniform image alone
does not prove that its example is useful.

Copy only reviewed captures into the source tree and update their hashes and
topic assignments together. Do not substitute unrelated historical renders.

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
