# Couverture automatisée du complément

## Validation de la fusion `codex/chat-ux` du 28 septembre 2026

La branche `72628f9` est intégrée à `main` après résolution des conflits avec les frontières de test, les gardes de session et les concepteurs WinForms existants. La mesure globale inclut toutes les sources de production apportées par cette branche, sans exclusion de production ni filtre de tests.

| Mesure | Résultat |
| --- | --- |
| Lignes | **17 211 / 18 517 — 92,94 %** |
| Branches | **17 509 / 19 944 — 87,79 %** |
| Classes instrumentées incomplètes | **63** |
| Suite complète, hôtes activés | **910 réussis, 0 échec, 0 ignoré** |
| Essais VSTest d’hôtes séparés | **2 Excel + 1 SOLIDWORKS réussis** |
| Concepteurs WinForms | **24 chargements DesignSurface et modifications de taille réussis** |
| Compilation du complément | **0 erreur, 0 avertissement** |
| Organisation des tests | **99 miroirs pour 168 fichiers de production** |

La couverture doit donc encore être complétée pour les nouvelles sources. La documentation de production à 100 % du bilan historique ne couvre pas automatiquement les nouvelles déclarations de cette branche.

### Preuves de la fusion

- Rapport : `artifacts/merge-chat-ux/final/f8fd3ea7-0be2-44d8-8e75-d1fa6b5ceb73/coverage.cobertura.xml`.
- Inventaire des classes incomplètes : `coverage-summary.json`, dans le même dossier.
- Suite globale : `artifacts/merge-chat-ux/final/global.trx`.
- Hôtes : `artifacts/merge-chat-ux/hosts/hosts.trx`.
- Concepteurs : `artifacts/merge-chat-ux/designers.txt`.
- Organisation : `artifacts/merge-chat-ux/final/test-layout.json`.

Excel est lancé visiblement par les fixtures, son VBE est ouvert par une commande native et les instances créées sont fermées après les essais. SOLIDWORKS 2019, version `27.5.0.0072`, est testé dans le PID fourni `23056`, conservé ouvert. L’identifiant de compilation de sa DLL chargée correspond à celui du fichier reconstruit.

Sur le projet jetable `test.swp`, les scripts de compilation/exécution et de breakpoint vérifient le pas à pas et la reprise, puis restaurent l’inventaire des modules et l’empreinte du code initial. La capture des variables locales indique un panneau accessible mais **zéro ligne exposée** ; elle ne prouve pas la lecture de leurs valeurs. Le diagnostic ROT externe reste `ProcessOnly`, tandis que le pont du complément et les commandes VBE fonctionnent dans ce PID.

Les matrices d’outils vérifient désormais `Rows` : types des cellules, dimensions, forme rectangulaire et limites de contenu. Les anciens tests d’interface sont adaptés aux onglets de paramètres et aux quatre colonnes fixes du diff. Le contrôle du catalogue conserve l’unicité et la longueur des noms de fonctions ; les capacités de taille du catalogue des fournisseurs distants ne sont pas qualifiées par ces simulations.

## Reproduire la mesure

Voir [le projet de tests](../../tests/README.md) pour les commandes VSTest et les hôtes opt-in. Les artefacts sont générés localement et ignorés par Git. Le seul exécutable exclu du collecteur est `ProviderTests.exe`, la simulation CLI ; aucune source du complément n’est exclue.

Le [bilan antérieur à 100 %](test-coverage-pre-chat-ux.md) est archivé séparément. Il ne couvre pas les nouvelles sources intégrées.
