# Licensing

VBAi uses the following licensing boundaries.

| Material | License | Where to find the terms |
| --- | --- | --- |
| Original VBAi source code: managed add-in, native renderer, updater, editor integration, scripts and tests | MPL-2.0 | [LICENSE](LICENSE) |
| Original VBA templates, examples and support runtimes intended to be copied into user projects | MIT | [LICENSES/MIT.txt](LICENSES/MIT.txt) |
| Third-party code, libraries, runtimes and assets | Their original licenses | [Third-party notices](THIRD_PARTY_NOTICES.md) and the accompanying upstream files |

## Original VBAi code

This Source Code Form is subject to the terms of the Mozilla Public License,
v. 2.0. If a copy of the MPL was not distributed with this file, you can obtain
one at [mozilla.org/MPL/2.0](https://www.mozilla.org/MPL/2.0/).

The MPL-2.0 is the default for original VBAi source code, except for the explicitly
identified MIT material and third-party components below. When distributing the
compiled add-in, make its corresponding MPL-covered sources available to recipients
under the terms of the license.

## VBA copied into user projects

The MIT grant covers original VBAi-authored VBA templates and examples offered for
reuse, the generated `VBAiTestSupport` module and the generated
`VBAiCoverageSupport` module. Their C# generators remain part of the MPL-2.0
add-in; the VBA they emit is licensed under MIT. The generated support modules
carry the complete MIT notice as VBA comments, so it accompanies exported modules.
Retain that notice when redistributing those modules or substantial portions.

Your own macro code and its repository keep the licensing terms you choose.
Using VBAi, adding its MIT support modules or exporting a macro to GitHub does
not place the entire macro project under the add-in's MPL-2.0 license.
Imported user code, third-party VBA and model-generated content are not
automatically relicensed by this policy; their applicable rights and terms remain
those of their authors and providers.

## Third-party components and distribution

Keep upstream copyright, license and notice files unchanged. The inventory and
provenance manifest in [third-party notices](THIRD_PARTY_NOTICES.md) identify the
bundled components. Installed software and external provider integrations retain
their vendors' own terms.

Application builds include `Licenses/MPL-2.0.txt`, `Licenses/MIT.txt` and
`Licenses/Scope.md`. Release payload preparation includes that folder together
with the bundled third-party notices. Selecting these licenses does not establish
an installer, a published release or vendor endorsement.
