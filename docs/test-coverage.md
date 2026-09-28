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

## Bilan historique avant la fusion `codex/chat-ux` du 28 septembre 2026

Cette mesure antérieure à la fusion porte sur l’assembly de production `CodexVBE`, avec la suite VSTest complète, sans filtre de tests ni exclusion de code de production. Elle couvre les lignes et les branches exécutables instrumentées par Coverlet ; elle ne constitue pas une preuve de compatibilité de chaque fonction avec tous les hôtes VBE. Les nouvelles sources apportées par `chat-ux` doivent être incluses dans une nouvelle mesure ; les 100 % ci-dessous ne leur sont pas attribués.

| Mesure | Résultat |
| --- | --- |
| Lignes | **15 472 / 15 472 — 100 %** |
| Branches | **15 610 / 15 610 — 100 %** |
| Classes instrumentées incomplètes | **0** |
| Tests de la suite complète | **826 réussis, 0 échec, 3 tests d’hôte désactivés** |
| Tests Excel activés séparément | **2 réussis, 0 échec, 0 ignoré** |
| Compilation du complément | **0 erreur, 0 avertissement** |

Le code mesuré correspond à `7133524`, également conservé par `c9a2389` : ce dernier rétablit uniquement les déclarations groupées de deux fixtures documentées. La comparaison Roslyn des sources avant et après le lot de documentation constate zéro différence de syntaxe exécutable, en production et dans les tests.

### Preuves locales

- Rapport global : `artifacts/coverage/complete-global-validated/b61106cb-d478-4943-8a28-e02a014a6948/coverage.cobertura.xml`.
- Résultats globaux : `artifacts/coverage/complete-global-validated/global.trx`.
- Résultats Excel : `artifacts/coverage/complete-excel-host/excel.trx`.
- Inventaire des miroirs : `artifacts/coverage/complete-global-validated/test-layout.json`.
- Comparaisons de documentation : `artifacts/coverage/complete-global/docs-production.txt` et `docs-tests.txt`.

Ces artefacts sont générés localement et ne sont pas ajoutés au dépôt.

### Commande de mesure

Depuis la racine du dépôt, dans PowerShell :

```powershell
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj -c Debug `
    '--collect:XPlat Code Coverage' `
    --results-directory artifacts/coverage/global `
    --logger 'trx;LogFileName=global.trx' `
    --blame-hang-timeout 4m -- `
    DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Format=cobertura `
    'DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Exclude=[ProviderTests]*'
```

`ProviderTests.exe` est un exécutable de simulation des processus fournisseurs. Cette exclusion ne concerne aucune source du complément. Aucun `ExcludeFromCodeCoverage`, `ExcludeByFile` ou `ExcludeByAttribute` n’est appliqué au code de production.

## Vérification dans Excel

La DLL et sa bibliothèque de types COM ont été reconstruites dans le chemin déjà enregistré, puis les tests ont été exécutés avec `CODEXVBE_RUN_EXCEL_TESTS=1` et le filtre `TestCategory=Excel`.

Les fixtures démarrent leurs propres instances Excel, affichent Excel et ouvrent le VBE par une commande native, sans raccourci ni coordonnées. Les tests vérifient le pont du complément, les projets, les fenêtres de débogage et la sauvegarde/relecture d’un classeur macro jetable. Ils ferment leurs classeurs et leurs processus après l’essai ; aucune instance Excel ne reste ouverte à l’issue de cette validation.

Le test SOLIDWORKS reste **NOT_RUN** : cet hôte n’était pas ouvert pour cette validation. Les doubles COM, les tests de bibliothèques COM Windows et la couverture instrumentée ne remplacent pas cet essai.

## Organisation et documentation

Le contrôle de rangement valide **85 fichiers de tests miroirs pour 96 fichiers de production**. Les fichiers Designer et certains contrats ou objets de données sont exercés par les scénarios complémentaires ; les interfaces COM et les métadonnées d’assembly n’ont pas de corps exécutable à tester.

La documentation IntelliSense de production est complète : **2 252 déclarations documentées, aucune manquante et aucune erreur de syntaxe**. La documentation des fixtures, des tests et des programmes de simulation est poursuivie séparément par l’agent dédié ; elle n’est pas déclarée complète dans ce bilan.

## Défauts corrigés pendant la validation

- Le statut de connexion Codex lit les deux flux du processus en parallèle pour que le délai puisse s’appliquer à un CLI bloqué.
- Une erreur de découverte des projets VBE est incluse dans le contexte LLM au lieu d’interrompre sa construction.
- La fixture Git utilise une entrée standard sans BOM, indépendante de l’ordre des tests. Le premier passage global a révélé cette dépendance ; le second passage complet, corrigé, est vert.
