# Development guide

Start with [source-build setup](installation.md), [architecture](architecture.md)
and [contributing](../CONTRIBUTING.md). Commands below run from the repository root.

## Builds and output directories

`Directory.Build.props` centralizes net48, C# 7.3 and x64. The normal add-in output
is `bin/<Configuration>/net48/`; the registration scripts currently use Debug.
Build an isolated output while any application holds the installed DLL:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet build VBAi.sln -c Release -p:BuildOutputRoot="$PWD/artifacts/build"
```

Each project then has its own `artifacts/build/<Project>/<Configuration>/net48/`
output. Do not use one shared `OutputPath` for all projects. An isolated build is
not automatically installed; verify the loaded assembly identity during native tests.

The native renderer build uses the Visual Studio C++ x64 tools and Windows SDK.
Generated native payloads and other test artifacts are not substitute source files.
See [native checks](../tests/native/README.md) for their qualification boundary.

## WinForms, WPF and runtime services

Keep `.cs`, `.Designer.cs` and `.resx` related with the proper project metadata.
Fixed controls belong in `InitializeComponent`; dynamic lists, document tabs and
transcript instances can be populated at runtime. Custom controls must keep their
properties editable and serializable in the Visual Studio Designer.

A Designer-compatible constructor must not start a host, open SQLite, create a
provider process or perform network I/O. DesignSurface checks should exercise
loading, property editing, resizing and serialization. They do not establish
correct live rendering, DPI behavior, focus or COM hosting.

The chat mixes WinForms views with WPF-hosted transcript/input surfaces; Monaco
uses WebView2. Preserve each component's ownership, threading and disposal rules.
Do not dispose or mutate a native parent reentrantly from a WebView callback.

Keep About/support metadata free of project contents. Runtime resource-reader
compatibility is handled locally; do not modify another application's configuration
to make a Designer or image resource load.

## Language services and generated editor assets

Edit `assets/editor/src/`, not minified files in `dist/`. Rebuild with:

```powershell
powershell.exe -NoProfile -File tools/Build-MonacoAssets.ps1
```

The lock files record versions and checksums. `-Offline` uses already-cached
archives; it does not magically supply a missing dependency. Commit the intended
generated distribution with its original license and notice files. Verify the
JavaScript editing/language tests and the WebView contracts described in
[testing](../tests/README.md).

## Localization and accessibility

Maintain the existing catalogs under `src/VBAi/Localization`. Preserve keys,
placeholders, formatting and technical identifiers. New visible text must be
localized using the project's existing mechanism rather than concatenating
untranslatable fragments.

Check long labels, keyboard use, tooltips, high contrast, DPI and right-to-left
layout where relevant. Catalog parity is not a linguistic review. Bundled Monaco
translations have their own coverage; do not claim that every widget is translated
just because the add-in has a corresponding UI language.

Product branding is VBAi; actual API names and migration identifiers must remain
literal. Review [icon provenance](../assets/icons/README.md) before replacing artwork.

## XML documentation

Maintain XML documentation for new or changed declarations, including relevant
private implementation contracts. The audit tool can inspect the add-in sources:

```powershell
dotnet run --project tools/XmlDocumentationAudit -- src/VBAi
```

Its `--compare <reference-source-directory>` mode checks syntax equivalence for
comment-only work. A documentation counter does not measure test coverage or the
quality of explanations. Avoid meaningless summaries that merely repeat a name.

## Documentation

Maintained Markdown is written in English and organized by reader task. Use the
README for the product overview, `docs/` for maintained guides, `tests/README.md`
for test commands, and root community files for contribution/support policies.

Keep one source of truth per subject. Update a guide in the same change as its
behavior. Link to code or discoverable schemas rather than hand-maintaining a
second exhaustive tool catalog. Put dated test evidence in `docs/test-coverage.md`
with the tested source/build and exact scope; do not repeat volatile counters in
the README, architecture or UI guides.

Do not add internal prompts, agent work logs, machine-specific worktree paths or
superseded experiments to the maintained guide tree. Use commits, issues, pull
requests and local artifacts for that history. Preserve relevant contracts before
removing an obsolete page. Check references from the solution and application
when moving entry points.

```sh
python -m unittest discover -s tools/docs -p "test_*.py"
python tools/docs/check_docs.py
```

The documentation checker requires Python 3.10 or later, has no third-party dependencies,
and checks maintained Markdown structure,
relative destinations and local Markdown anchors. It does not certify external
link availability or technical accuracy. Run product/host tests only when the
change warrants them, and report exactly what was and was not executed.
