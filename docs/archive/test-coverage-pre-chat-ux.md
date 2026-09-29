# Couverture avant la fusion chat-ux

> Bilan historique. La mesure actuelle figure dans [test-coverage.md](../test-coverage.md).


Cette mesure antérieure à la fusion porte sur l’assembly de production `VBAi`, avec la suite VSTest complète, sans filtre de tests ni exclusion de code de production. Elle couvre les lignes et les branches exécutables instrumentées par Coverlet ; elle ne constitue pas une preuve de compatibilité de chaque fonction avec tous les hôtes VBE. Les nouvelles sources apportées par `chat-ux` doivent être incluses dans une nouvelle mesure ; les 100 % ci-dessous ne leur sont pas attribués.

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
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug `
    '--collect:XPlat Code Coverage' `
    --results-directory artifacts/coverage/global `
    --logger 'trx;LogFileName=global.trx' `
    --blame-hang-timeout 4m -- `
    DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Format=cobertura `
    'DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Exclude=[ProviderTests]*'
```

`ProviderTests.exe` est un exécutable de simulation des processus fournisseurs. Cette exclusion ne concerne aucune source du complément. Aucun `ExcludeFromCodeCoverage`, `ExcludeByFile` ou `ExcludeByAttribute` n’est appliqué au code de production.

## Vérification dans Excel

La DLL et sa bibliothèque de types COM ont été reconstruites dans le chemin déjà enregistré, puis les tests ont été exécutés avec `VBAi_RUN_EXCEL_TESTS=1` et le filtre `TestCategory=Excel`.

Les fixtures démarrent leurs propres instances Excel, affichent Excel et ouvrent le VBE par une commande native, sans raccourci ni coordonnées. Les tests vérifient le pont du complément, les projets, les fenêtres de débogage et la sauvegarde/relecture d’un classeur macro jetable. Ils ferment leurs classeurs et leurs processus après l’essai ; aucune instance Excel ne reste ouverte à l’issue de cette validation.

Le test SOLIDWORKS reste **NOT_RUN** : cet hôte n’était pas ouvert pour cette validation. Les doubles COM, les tests de bibliothèques COM Windows et la couverture instrumentée ne remplacent pas cet essai.

## Organisation et documentation

Le contrôle de rangement valide **85 fichiers de tests miroirs pour 96 fichiers de production**. Les fichiers Designer et certains contrats ou objets de données sont exercés par les scénarios complémentaires ; les interfaces COM et les métadonnées d’assembly n’ont pas de corps exécutable à tester.

La documentation IntelliSense de production est complète : **2 252 déclarations documentées, aucune manquante et aucune erreur de syntaxe**. La documentation des fixtures, des tests et des programmes de simulation est poursuivie séparément par l’agent dédié ; elle n’est pas déclarée complète dans ce bilan.

## Défauts corrigés pendant la validation

- Le statut de connexion Codex lit les deux flux du processus en parallèle pour que le délai puisse s’appliquer à un CLI bloqué.
- Une erreur de découverte des projets VBE est incluse dans le contexte LLM au lieu d’interrompre sa construction.
- La fixture Git utilise une entrée standard sans BOM, indépendante de l’ordre des tests. Le premier passage global a révélé cette dépendance ; le second passage complet, corrigé, est vert.
