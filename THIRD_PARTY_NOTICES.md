# Third-party components and notices

VBAi uses third-party software and identifies integrations with vendor trademarks.
Those components retain their own licenses and notices. This inventory is not a
license for VBAi and is not a substitute for the terms distributed with each
component.

## Bundled and referenced components

| Component | Source of version and notices |
| --- | --- |
| Monaco Editor and bundled editor resources | `assets/editor/monaco.lock.json`; [Monaco license](assets/editor/dist/MONACO-LICENSE.txt) and [third-party notices](assets/editor/dist/MONACO-ThirdPartyNotices.txt). |
| Microsoft WebView2 | Package reference in `src/VBAi/VBAi.csproj`; Microsoft runtime distribution terms also apply. The loader and the installed runtime are distinct components. |
| Markdig | Package reference in `src/VBAi/VBAi.csproj`; preserve the package's upstream license and notices. |
| System.Resources.Extensions | Package reference in `src/VBAi/VBAi.csproj`; preserve the package's upstream license and notices. |
| esbuild | `assets/editor/esbuild.lock.json`; used to build editor assets. |
| Test and development dependencies | Test project files and the tooling project files record their dependencies. |

Do not remove upstream license files from generated assets. When preparing a
release, inventory the resolved dependency closure as well as direct references,
and include the notices required by the actual payload. This table does not claim
that a full release-license audit has been completed.

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
license and a review of redistributed components must be handled explicitly by
the maintainer; this documentation update does not choose or change licensing terms.
