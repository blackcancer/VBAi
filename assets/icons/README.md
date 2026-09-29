# Icon assets

VBAi uses the following source artwork for its Windows forms and project identity:

| File | Origin and purpose |
| --- | --- |
| `assistant.png` | Generated VB monogram with a violet AI accent on a graphite tile; VBAi product identity. |
| `settings.png` | Generated adjustment sliders in the same visual family. |
| `github.png` | Unmodified GitHub Invertocat Black Clearspace from the official GitHub brand toolkit, retrieved on September 27, 2026. |

The GitHub mark identifies the GitHub integration. It remains GitHub's trademark
and does not imply endorsement. Consult the [GitHub brand toolkit](https://brand.github.com/foundations/logo)
and [project notices](../../THIRD_PARTY_NOTICES.md) before redistributing or changing it.

## Regenerate Windows icons

Run from the repository root:

```powershell
powershell.exe -NoProfile -File tools/Build-WindowIcons.ps1
```

ICO files contain 16, 20, 24, 32, 48, 64, 128 and 256 pixel frames. The script also
updates embedded form icon resources. Keep each form's `Icon` assignment in its
Designer so the property remains editable in Visual Studio. Runtime forms must
not depend on paths to the source PNG files.

Product-facing text uses **VBAi — Your AI agent for VBA**. Do not rename persistent
COM GUIDs or migration identifiers as an artwork change. Use the current
[source-build instructions](../../docs/installation.md); historical resource-compiler
workarounds are not a separate build procedure.
