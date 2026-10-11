# Public website

The bilingual product website is published at
[vbai.app](https://vbai.app/), hosted on GitHub Pages.
French pages live under `/fr/`; the root is English. Each page has a distinct
title and description, a canonical URL and reciprocal language alternatives.
The generated sitemap contains the eight localized product pages.

The homepage presents AI assistance and GitHub integration as the two main product
capabilities. Its GitHub section follows the opening scene, with the real exported
source view, benefits for VBA projects and links to the macro-to-repository walkthrough.
The internal "GitHub for VBA" navigation is separate from the external source-code
link. Keep the explanation aligned with `docs/github-integration.md`: document
saving, local commits, remote pushes and reviewed source imports are separate actions.

## Sources and build

Manual English and French text lives in `content/en.json` and `content/fr.json`.
The shared template, stylesheet and optional screenshot viewer are in this folder.
The builder uses Python 3.10 or newer and no third-party dependencies:

```sh
python tools/docs/build_site.py --output artifacts/site-preview/VBAi
python -m http.server 8765 --directory artifacts/site-preview/VBAi
```

Open `http://localhost:8765/` or `/fr/`. Serve the generated output folder
so the root-relative URLs behave exactly as they do on the custom domain.
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
The original graphite/violet logo defines its visual identity. The opening copy
enters in sequence; the two real captures float independently, and the violet
spark rotates and breathes. Pointer tilt and staggered scroll reveals complete
the motion. Floating captures pause on hover or keyboard focus; the scene pauses
off screen and decorative motion pauses in a hidden tab. Reduced motion disables
all these effects, including when that preference changes while the page is open.
Navigation/content remain available without JavaScript. CSS and JavaScript URLs
include content hashes so an update does not keep an older asset in browser cache.
Its documentation links point to the authoritative repository guides rather
than maintaining a second copy. The generated `robots.txt` advertises the sitemap
at the custom-domain root.

## Google Search Console

To monitor indexing and traffic from Google Search:

1. Sign in to [Google Search Console](https://search.google.com/search-console/)
   with the maintainer's Google account. Add the URL-prefix property
   `https://vbai.app/`, or select it if it already exists.
2. If ownership is not already verified, select **HTML tag** and copy the exact
   `content` value from Google's `google-site-verification` meta tag into
   `google_site_verification` in `config.json`. Leave the setting empty when no
   token has been supplied; the builder then emits no verification tag.
3. Build the website and deploy the reviewed change through the Pages workflow.
   Check that the published homepage contains the exact tag in its HTML `head`,
   then click **Verify** in Search Console. Keep the token in the configuration:
   Google periodically checks that it remains present. The token is public HTML
   metadata; do not put Google credentials or OAuth access tokens in this setting.
4. Submit `https://vbai.app/sitemap.xml` in **Sitemaps**. It contains the eight
   localized product pages; successful submission does not prove indexing.
5. Inspect `https://vbai.app/` and `https://vbai.app/fr/` with **URL inspection**.
   Check the live URL and request indexing if needed. Use **Page indexing** to
   track Google's actual indexing status and exclusion reasons.
6. Use **Performance** to review impressions, clicks, queries and landing pages.
   These reports cover Google Search, not all visits to the website, and new
   properties can take a few days to accumulate data.

Google account sign-in and ownership verification require access to the
maintainer's account. A local build or successful public HTTP check does not
establish that the property is verified, the sitemap is submitted or a URL is
indexed. See Google's [ownership verification guidance](https://support.google.com/webmasters/answer/9008080)
and [recrawl guidance](https://developers.google.com/search/docs/crawling-indexing/ask-google-to-recrawl).

## Domain configuration

`vbai.app` is the custom domain configured in the repository's GitHub Pages
settings. The OVH DNS zone must point the apex to GitHub Pages with these four
`A` records: `185.199.108.153`, `185.199.109.153`, `185.199.110.153` and
`185.199.111.153`. `www` uses a `CNAME` to `blackcancer.github.io.`.
Preserve mail and other unrelated records; replace conflicting web records only.
The main domain is `vbai.app`; GitHub Pages redirects the configured `www` variant.
HTTPS enforcement is enabled once GitHub has issued the domain certificate.
This deployment uses Actions, so Pages stores the domain in its settings;
a source `CNAME` file is not required. Changing the domain also requires updating
`config.json` and rebuilding canonical URLs, language alternatives and the sitemap.
