# Offline UI translation tools

The application ships only embedded `.resx` catalogues. It never runs a translation
model or sends UI strings to a translation service.

Development tooling: Python 3.11, ctranslate2 4.8.2, sentencepiece 0.2.2,
opencc-python-reimplemented 0.1.7. Install these in an isolated virtual environment,
not in the application or system Python.

Download model packages separately from the public
[Argos model index](https://github.com/argosopentech/argospm-index).
Only public model files are downloaded; project strings stay local.
Extract each package into `artifacts/localization/models/<code>/<package>/`, keeping
`sentencepiece.model`, `metadata.json` and `model/`. Stanza is not required for these
short UI strings. Used English-source package versions: ar 1.0, de 1.3, es 1.0,
hi 1.1, it 1.0, ja 1.1, ko 1.1, pt 1.9, ru 1.9, zh 1.9.

```powershell
artifacts/local-translation/Scripts/python.exe tools/localization/translate_offline.py de
# Generate simplified Chinese before traditional Chinese.
artifacts/local-translation/Scripts/python.exe tools/localization/translate_offline.py zh
artifacts/local-translation/Scripts/python.exe tools/localization/translate_offline.py zt
```

Existing translations are preserved; `overrides.json` always takes precedence.
Delete only a specific resource entry to request its regeneration, or edit the resource
directly. Add reviewed corrections to `overrides.json` to make them reproducible.
Do not ship raw model output without checking punctuation, repeated output, technical
identifiers, negation and UI layout. These catalogues still need full native-language review.

Local reports go to `artifacts/localization/audit-<code>.json`. Model files, Python
packages and reports are ignored by Git. The traditional Chinese model was rejected
because of repetitions; `zt` uses the local OpenCC `s2twp` conversion with overrides.

References: [CTranslate2](https://opennmt.net/CTranslate2/),
[Argos tokenizer](https://github.com/argosopentech/argos-translate/blob/master/argostranslate/tokenizer.py),
[OpenCC](https://github.com/BYVoid/OpenCC).
