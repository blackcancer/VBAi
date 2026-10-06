# Third-party components and notices

VBAi uses third-party software and identifies integrations with vendor trademarks.
Those components retain their own licenses and notices. This inventory is not a
replacement for their upstream terms. Original VBAi code and reusable VBA support
have separate grants described in [licensing scope](LICENSING.md).

## Bundled runtime components

The runtime package references are declared in `src/VBAi/VBAi.csproj`; NuGet resolves
their dependency closure in `src/VBAi/obj/project.assets.json`. The
[notice manifest](assets/third-party/manifest.json) records the inspected package
versions, package digests, upstream provenance and notice-file SHA-256 values.
Upstream license and notice files are preserved verbatim.

| Component | Version | Included terms and notices |
| --- | --- | --- |
| Markdig | 0.37.0 | [BSD-2-Clause license](assets/third-party/Markdig-0.37.0/LICENSE.txt), retrieved from the exact repository commit recorded by the package's `.nuspec`. |
| Microsoft WebView2 SDK assemblies and native loader | 1.0.4258.31 | Package [license](assets/third-party/Microsoft.Web.WebView2-1.0.4258.31/LICENSE.txt) and [notices](assets/third-party/Microsoft.Web.WebView2-1.0.4258.31/NOTICE.txt). These describe the SDK package, not a bundled browser runtime. |
| System.Resources.Extensions | 8.0.0 | [MIT license](assets/third-party/System.Resources.Extensions-8.0.0/LICENSE.TXT) and [package notices](assets/third-party/System.Resources.Extensions-8.0.0/THIRD-PARTY-NOTICES.TXT). |
| System.Memory | 4.5.5 | [MIT license](assets/third-party/System.Memory-4.5.5/LICENSE.TXT) and [package notices](assets/third-party/System.Memory-4.5.5/THIRD-PARTY-NOTICES.TXT). |
| System.Buffers | 4.5.1 | [MIT license](assets/third-party/System.Buffers-4.5.1/LICENSE.TXT) and [package notices](assets/third-party/System.Buffers-4.5.1/THIRD-PARTY-NOTICES.TXT). |
| System.Numerics.Vectors | 4.5.0 | [MIT license](assets/third-party/System.Numerics.Vectors-4.5.0/LICENSE.TXT) and [package notices](assets/third-party/System.Numerics.Vectors-4.5.0/THIRD-PARTY-NOTICES.TXT). |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | [MIT license](assets/third-party/System.Runtime.CompilerServices.Unsafe-4.5.3/LICENSE.TXT) and [package notices](assets/third-party/System.Runtime.CompilerServices.Unsafe-4.5.3/THIRD-PARTY-NOTICES.TXT). |
| Monaco Editor, font, translations and bundled editor resources | 0.55.1 | [Locked package](assets/editor/monaco.lock.json), [MIT license](assets/editor/dist/MONACO-LICENSE.txt) and [bundled third-party notices](assets/editor/dist/MONACO-ThirdPartyNotices.txt). |

The add-in project copies `assets/third-party` into `ThirdPartyNotices` beside the
assembly and the Monaco notices into `EditorAssets`. Keep both directories with
the complete output when distributing it. Update the provenance manifest and
upstream texts whenever a dependency changes; a prior notice bundle is not evidence
for a different version. Review the actual final payload, including transitive
packages, before a release. Do not remove notices from generated editor assets.

## Native code, prerequisites and build tools

`VBAi.Native.dll` is built from this repository and embedded in `VBAi.dll`.
[The native build](tools/build/Build-NativeRenderer.ps1) selects `/MT`, incorporating
Microsoft C/C++ runtime code instead of shipping a separate runtime DLL. The
selected compiler, standard-library and Windows SDK terms remain relevant to the
resulting binary; static linking does not establish distribution rights by itself.
The Microsoft standard-library headers identify their
[Apache-2.0 license with LLVM exception](https://github.com/microsoft/STL/blob/main/LICENSE.txt).
Consult the redistribution terms for the actual Visual Studio edition/toolset used
for the release, such as the [Visual Studio 2026 distribution list](https://learn.microsoft.com/en-us/visualstudio/releases/2026/redistribution).

Windows `winsqlite3.dll`, .NET Framework 4.8, Office/SOLIDWORKS and the Evergreen
WebView2 browser runtime are external host/runtime prerequisites, not binaries
included by these package-content rules. The WebView2 SDK loader and installed
browser runtime are separate components. A future runtime bootstrapper or installer
payload needs its own distribution review; this notice inventory adds neither.

[esbuild 0.25.12](assets/editor/esbuild.lock.json) is an editor build tool. Its
executable is not part of the add-in's content items. Test SDKs, coverage collectors
and diagnostic fixtures are development dependencies, not intended release files.
If build tools or test outputs are distributed separately, inventory their own
resolved closure and preserve their upstream terms as well.

## Icons and trademarks

The GitHub icon is the unmodified Invertocat asset identified in the
[icon provenance guide](assets/icons/README.md). It identifies the GitHub integration
and remains GitHub's trademark. Product artwork and command icons must not be
represented as a grant of rights to third-party brands.

Names such as Microsoft, Visual Basic, Excel, SOLIDWORKS, GitHub, OpenAI and other
provider names identify external products or integrations. Their appearance does
not imply endorsement or affiliation.

## VBAi licensing

The repository does not yet contain a selected license for VBAi itself. A project
license must be handled explicitly by the maintainer. Copying upstream notices does
not select a VBAi license or constitute legal clearance for a release.
