# Couverture automatisée du complément

## Fusion des extensions `codex/chat-ux` du 28 septembre 2026

Fusion locale `611dcb5`, intégrant `5aed4f3` dans main après l’isolation des sessions CLI. Compilation solution et complément : **0 erreur, 0 avertissement**. Les sources CLI privées sont conservées.

| Mesure | Résultat |
| --- | --- |
| Lignes | **19 201 / 19 299 — 99,49 %** |
| Branches | **20 962 / 21 297 — 98,43 %** |
| Classes instrumentées incomplètes | **19** |
| Tests locaux, hors trois essais hôtes | **1 085 réussis** |
| Suite globale avec hôtes | **1 086 réussis, 1 échec, 1 ignoré** |
| Excel, lecture et pont | Réussi dans la suite globale |
| Excel, sauvegarde | Échec dans la suite globale ; réussi dans le passage isolé |
| SOLIDWORKS | NOT_RUN : instance utilisateur fermée |
| Organisation | **128 miroirs pour 177 sources**, scénarios complémentaires |

### Diagnostic Excel

Le test de sauvegarde échoue avec « The registered Excel instance is not this VBE host. ». Le contrôle de PID de production empêche la sauvegarde dans une autre instance retournée par le registre COM. Le même test réussit séparément. La suite globale avec hôtes n’est donc pas entièrement verte : l’enchaînement/la cohabitation d’instances Excel utilisant ce registre reste à stabiliser. Aucun contrôle de sécurité ni assertion de sauvegarde n’a été retiré ; l’assertion rapporte maintenant la raison native.

### Preuves

- Rapport global : `artifacts/merge-chat-ux-latest/final/541bf7ad-5bdd-41fc-a733-18611912a750/coverage.cobertura.xml`.
- Inventaire : `coverage-summary.json`, à côté du rapport ; les nouvelles déclarations, mutations et frontières natives expliquent les lacunes.
- Suite globale : `artifacts/merge-chat-ux-latest/final/global.trx`.
- Sauvegarde Excel isolée : `artifacts/merge-chat-ux-latest/hosts-diagnostic/hosts.trx`.
- Catalogue reconstruit : **177 outils**, documentation vérifiée sans divergence.
- Convention miroir : `tools/tests/Test-TestLayout.ps1`.

Les [extensions fonctionnelles](reference/functional-extensions.md) détaillent les nouveaux contrats et leurs qualifications natives. Le [bilan précédent à 100 % après isolation CLI](archive/test-coverage-session-isolation.md) ne couvre pas ces ajouts. Les 24 DesignSurface et les essais SOLIDWORKS précédents restent historiques, non réexécutés ici. L’audit IntelliSense continue dans sa branche dédiée.

## Reproduire la mesure

Voir [le projet de tests](../tests/README.md) pour VSTest, la couverture et les hôtes opt-in. Les artefacts sont locaux et ignorés par Git. `ProviderTests.exe` est le seul exécutable exclu ; aucune source du complément n’est exclue. Une couverture VSTest comprend les doubles natifs et ne qualifie pas toutes les commandes dans chaque hôte.
