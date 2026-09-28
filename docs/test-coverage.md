# Couverture automatisée du complément

## Mesure globale après intégration de la PR #4 et du lot IDE

Mesure du **28 septembre 2026**, code **`27389a8`**, comprenant la PR WinForms `codex/chat-ux`, le renommage de paramètres, les appels nommés et le vérificateur de signatures. Compilation de la solution : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Lignes exécutables | **20 639 / 20 639 — 100 %** |
| Branches | **22 223 / 22 223 — 100 %** |
| Classes instrumentées incomplètes | **0** |
| Suite globale VSTest | **1 172 réussis, 0 échec, 1 ignoré** ; 6 min 37 s |
| Excel automatisé | **Trois scénarios réussis** dans la suite globale |
| SOLIDWORKS | **NOT_RUN** : aucune instance utilisateur préchargée |
| Organisation | **141 miroirs pour 193 sources**, scénarios et fixtures complémentaires |
| Concepteurs WinForms | **27 DesignSurface** chargées, redimensionnées et éditées ; contrôles enfants sélectionnables |
| Catalogue LLM | **180 outils** extraits de l'assembly reconstruite |

## Scénarios du lot

- Renommage de paramètre privé standard : déclaration, usages, appels nommés directs/qualifiés et imbriqués, continuations, collisions, compilation conditionnelle, SHA/mode, diff LLM et undo/redo.
- Exécution publique avec noms de paramètres vérifiés, réordonnancement et omission d'optionnels ; requêtes/noms/valeurs capturés avant la livraison ; gardes de taille, type, identité et syntaxe.
- Vérificateur Office : formats, disponibilité du SIP VBA, SHA et verrou de lecture, résultats HRESULT distincts, relecture du registre sans écriture et nettoyage WinVerifyTrust. Les résultats cryptographiques sont simulés localement ; leur qualification native demeure NOT_RUN tant que le SIP dédié n'est pas autorisé.
- PR WinForms : propriétés et marges du moteur de saisie, sérialisation/reset, apparence, événements avec/sans abonnés, ressources Designer, et libération des diff natifs lors du recyclage du transcript.
- Contrat du catalogue LLM : tous les champs requis/facultatifs, types et modes ; `ArgumentNames` validé comme tableau, sans modifier les politiques d'approbation.

Excel vérifie indépendamment les cellules **42** et **named call** après renommage et appel nommé, puis la restauration exacte du code. Le document enregistré est copié pour la vérification : Excel garde un verrou d'écriture sur son fichier ouvert. Le SIP absent produit **VerifierUnavailable**, pas une signature valide. Les deux autres scénarios vérifient le chargement/pont et la sauvegarde en présence d'une autre instance Excel jetable, laissée inchangée.

## Preuves et limites

- Rapport global : `artifacts/pr4-integration/final-qualified/a7899ebd-1a42-41ce-9820-9ebf2faa2aa3/coverage.cobertura.xml` ; résumé `coverage-summary.json` adjacent.
- Résultats globaux : `artifacts/pr4-integration/final-qualified/global.trx`.
- Organisation miroir : `artifacts/pr4-integration/test-layout.json`, produite par `tools/tests/Test-TestLayout.ps1`.
- Scripts Designer : `tools/tests/Test-WinFormsDesigners.ps1` et `tools/tests/Test-ChatDesigner.ps1`, exécutés sur l'assembly de la fusion.
- [Mesure précédente à 100 %](archive/test-coverage-pre-winforms-pr4.md) : historique, distinct de ce passage.

Aucune source du complément n'est exclue ; `ProviderTests.exe` est le seul exécutable de simulation exclu. Le collecteur mesure le processus VSTest, pas le code exécuté dans Excel ou SOLIDWORKS. Les doubles natifs vérifient les branches et erreurs du code ; 100 % de couverture ne qualifie pas toutes les fonctions dans tous les hôtes.

L'[inventaire fonctionnel complet du VBE](reference/vbe-capability-inventory.md) précise les fonctions absentes ou partielles. Les [qualifications natives](reference/native-qualification.md) distinguent les preuves Excel, les essais historiques SOLIDWORKS et le digest VBA non qualifié. L'audit IntelliSense avec Luna reste indépendant et non intégré à cette mesure.

## Reproduire la mesure

Voir [le projet de tests](../tests/README.md) pour VSTest, la couverture et les hôtes opt-in. Les rapports détaillés restent locaux sous `artifacts/`, ignoré par Git.
