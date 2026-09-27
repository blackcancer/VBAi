# Window icons

- `assistant.png`: generated VB monogram and violet AI sparkle on a graphite tile.
- `settings.png`: generated matching adjustment sliders.
- `github.png`: unmodified GitHub Invertocat Black Clearspace from the official
  [GitHub brand toolkit](https://brand.github.com/foundations/logo), downloaded
  from `https://brand.github.com/GitHub_Logos.zip` on 2026-09-27.
  GitHub's mark identifies the GitHub integration only. It remains GitHub's trademark.

The ICO files contain 16, 20, 24, 32, 48, 64, 128 and 256 pixel frames.
Run `powershell -File tools/Build-WindowIcons.ps1` to regenerate the ICO files
and the embedded `$this.Icon` form resources from the PNG sources.
Each form's Icon property is assigned in InitializeComponent and remains editable
in the Visual Studio WinForms designer. No runtime asset path is required.

Product identity: **VBAi — Your AI agent for VBA**.
COM identifiers and existing data paths retain their original names for compatibility.

Build with Visual Studio / .NET Framework MSBuild. The SDK-only `dotnet build`
resource compiler cannot serialize WinForms Icon resources without additional
runtime dependencies; the Visual Studio build uses the native framework resource compiler.
