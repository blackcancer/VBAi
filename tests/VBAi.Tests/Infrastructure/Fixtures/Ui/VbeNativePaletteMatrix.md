# Palette recovery and transaction matrix

Prepared before test execution. All paths are disposable temporary directories; dialog visits use a scoped in-memory palette. No Office process or real VBE settings.

| Area | Cases and asserted behavior |
| --- | --- |
| Row validation | null/wrong count/null row/blank name/duplicate name; each lower/upper color bound; inclusive bounds accepted |
| Equality | null sides, length, null rows, each independent field mismatch, equal empty/full palettes |
| Recovery JSON | missing, oversized, null, malformed, schema/version mismatch, invalid original/applied, inconsistent dark map; successful roundtrip and failed overwrite preserve bytes |
| Transaction | restore without state performs no visits; first apply snapshots original; repeat apply performs no update; restore verifies reopened state then deletes snapshot |
| Failure | current manually changed, reopened palette mismatch, dialog failure, concurrent lock; retain original snapshot and release lock |
| Map | exact ten foreground/background values, preserved names/indicators, input untouched |

Production transaction uses the same file lock, recovery validation, callback, reopen comparison and deletion. The overload accepts only the version and dialog visit boundary; the host overload supplies the real dialog implementation.
