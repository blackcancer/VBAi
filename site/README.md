# Public website

The bilingual product website is published at
[blackcancer.github.io/VBAi](https://blackcancer.github.io/VBAi/).
French pages live under `/fr/`; the root is English. Each page has a distinct
title and description, a canonical URL and reciprocal language alternatives.
The generated sitemap contains the eight localized product pages.

## Sources and build

Manual English and French text lives in `content/en.json` and `content/fr.json`.
The shared template, stylesheet and optional screenshot viewer are in this folder.
The builder uses Python 3.10 or newer and no third-party dependencies:

```sh
python tools/docs/build_site.py --output artifacts/site-preview/VBAi
python -m http.server 8765 --directory artifacts/site-preview
```

Open `http://localhost:8765/VBAi/` or `/VBAi/fr/`. Serve the parent output folder
so the project-relative URLs behave exactly as they do on GitHub Pages.
Generated files are ignored and must not be committed.

## Screenshots and publication

`config.json` lists the published screenshots. The builder copies their original
pixels from the [help captures](../docs/help/fr-FR/screenshots/manifest.json)
only after validating live provenance and the exact reviewed SHA-256. It does
not publish that manifest, local test receipts, CHM staging or application binaries.
The product icon comes from the existing application assets. Captions explain
the real disposable CalculTVA scenario; French screenshots are identified on
the English site. No mock UI, invented reviews or acceptance badges are used.
Reviewed CSS framing removes the black Windows capture margins in both the
page and the enlarged viewer. The image files and interface pixels remain original.

The [Pages workflow](../.github/workflows/pages.yml) builds and checks local links,
anchors and image alternatives for pull requests; only main/manual runs deploy.
Deployment uses the GitHub Pages workflow source and uploads only the dedicated
generated folder. Update the version and verified download URL together in
`config.json`, and update the corresponding release wording in both languages.

The website has no analytics, cookies, remote fonts or third-party JavaScript.
The original graphite/violet logo defines its visual identity. Screenshot tilt,
the decorative spark and scroll reveals are progressive enhancements; reduced
motion disables them, and navigation/content remain available without JavaScript.
Its documentation links point to the authoritative repository guides rather
than maintaining a second copy. Submit the published `sitemap.xml` through a
verified Google Search Console URL-prefix property to monitor indexing.
Search Console verification remains a maintainer account operation. A project
site cannot control `robots.txt` at the `blackcancer.github.io` domain root.
