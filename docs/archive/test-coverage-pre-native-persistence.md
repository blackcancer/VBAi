# Couverture avant la PR de persistance native

Mesure globale du 28 septembre 2026, code **`f35da60`**, avant intégration de la PR #2 (`codex/chat-ux`).

| Mesure | Résultat |
| --- | --- |
| Lignes exécutables | **19 354 / 19 354 — 100 %** |
| Branches | **21 321 / 21 321 — 100 %** |
| Classes instrumentées incomplètes | **0** |
| Tests réussis | **1 123** |
| Échecs | **0** |
| SOLIDWORKS | **1 scénario ignoré**, aucune instance préchargée |

Les deux scénarios Excel passent dans cette suite. Le test de sauvegarde ouvre deux instances jetables distinctes et vérifie que l'autre classeur reste non enregistré. La résolution de l'application utilise le document natif `EXCEL7` du processus hôte, avec repli ROT conservant le contrôle de PID.

Rapport local : `artifacts/coverage-completion/final/4b833348-4af7-4dfa-9aae-24c2329f91be/coverage.cobertura.xml`, inventaire `coverage-summary.json` adjacent et résultats `artifacts/coverage-completion/final/global.trx`.

Aucune source de production n'est exclue. Seul l'exécutable de simulation `ProviderTests.exe` est exclu de la mesure. Ce bilan ne couvre pas les ajouts ultérieurs de la PR #2 et ne qualifie pas toutes les fonctions dans tous les hôtes.
