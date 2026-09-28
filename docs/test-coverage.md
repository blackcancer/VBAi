# Couverture automatisée du complément

## Mesure globale du 28 septembre 2026

Code testé : `fa61565`, après intégration de `codex/chat-ux` et achèvement des matrices de l’éditeur, du concepteur, du débogueur et de la conversation. Toute l’assembly de production est instrumentée, sans exclusion de production ni filtre de tests.

| Mesure | Résultat |
| --- | --- |
| Lignes | **18 537 / 18 537 — 100 %** |
| Branches | **19 936 / 19 936 — 100 %** |
| Classes instrumentées incomplètes | **0 sur 352** |
| Suite complète | **1 067 réussis, 0 échec, 1 ignoré** |
| Hôtes Excel | **2 essais réussis**, instances visibles isolées puis fermées |
| Hôte SOLIDWORKS | **NOT_RUN** dans ce passage : instance utilisateur fermée |
| Compilation du complément | **0 erreur, 0 avertissement** |
| Organisation des tests | **119 miroirs pour 168 fichiers de production**, scénarios complémentaires séparés |

### Preuves

- Rapport : `artifacts/coverage-resume/global-final/32e372fa-ccab-4261-98ac-2a3c0c17f376/coverage.cobertura.xml`.
- Inventaire : `coverage-summary.json`, dans le même dossier ; aucune classe incomplète.
- Suite globale : `artifacts/coverage-resume/global-final/global.trx`.
- Contrôle miroir : `tools/tests/Test-TestLayout.ps1`.

Les assertions des matrices vérifient les lectures, modifications, versions, gardes, résultats partiels et restaurations exactes. Les frontières natives injectables conservent les implémentations Windows/COM par défaut. Les méthodes et branches de production ne sont pas masquées pour améliorer les chiffres.

### Portée native

Excel est lancé visiblement par les fixtures et son VBE est ouvert avec une commande native. Les tests inspectent le complément connecté et vérifient sauvegarde/relecture d’un classeur macro jetable ; ils ferment uniquement leurs propres instances. La DLL installée a été reconstruite avant ces essais.

La couverture du processus VSTest comprend des doubles COM et des simulations de frontières natives. Elle ne prouve pas chaque combinaison de commandes dans tous les hôtes, ni les fournisseurs authentifiés. Les essais SOLIDWORKS et les 24 chargements DesignSurface précédents sont conservés dans le [bilan daté de la fusion](archive/test-coverage-chat-ux-integration.md) ; ils n’ont pas été réexécutés dans ce passage. La documentation IntelliSense continue dans sa branche dédiée et reste à auditer après intégration.

## Reproduire la mesure

Voir [le projet de tests](../tests/README.md) pour VSTest, la couverture et les hôtes opt-in. Les artefacts sont locaux et ignorés par Git. Le seul exécutable exclu du collecteur est `ProviderTests.exe`, la simulation CLI ; aucune source du complément n’est exclue.

Le [bilan antérieur à chat-ux](archive/test-coverage-pre-chat-ux.md) reste archivé séparément.
