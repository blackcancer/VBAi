# Couverture automatisée du complément

## Mesure globale après intégration des PR #2 et #3

Mesure du **28 septembre 2026**, code **`fb166a4`**, comprenant la persistance native de `codex/chat-ux`, la récupération du placement du panneau et les tests complémentaires. Compilation du complément : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Lignes exécutables | **19 546 / 19 546 — 100 %** |
| Branches | **21 793 / 21 793 — 100 %** |
| Classes instrumentées incomplètes | **0** |
| Suite globale VSTest | **1 151 réussis, 0 échec, 1 ignoré** |
| Excel automatisé | **Deux scénarios réussis** dans la suite globale |
| SOLIDWORKS | **NOT_RUN** : aucune instance utilisateur préchargée |
| Organisation | **134 miroirs pour 183 sources**, scénarios et fixtures complémentaires |

## Scénarios ajoutés

- Profils SQLite des barres d'outils : transactions, validation, restauration, commandes temporaires et persistantes, collisions et erreurs natives. Une suppression refusée conserve le profil.
- Concepteur : dispatch typé des propriétés scalaires natives et repli sur les descripteurs, avec erreurs de conversion.
- Options : sélection d'un contrôle éditable unique malgré des libellés homonymes, refus d'un libellé seul, vérification de relecture et annulation.
- Débogueur : regroupement des observations UIA identiques ; valeurs et ascendants distincts restent séparés.
- Hôte : initialisation des profils Excel/SOLIDWORKS, nettoyage à la déconnexion, dimensions utilisables préservées et récupération des panneaux trop petits.

Le test Excel de sauvegarde crée deux instances jetables distinctes. Il vérifie la sauvegarde du projet visé et confirme que l'autre classeur reste non enregistré. La résolution passe par le document natif `EXCEL7` du processus propriétaire, avec le repli ROT et les contrôles de PID conservés. Les tests ferment leurs propres instances uniquement.

## Preuves et limites

- Rapport global : `artifacts/coverage-post-prs/final/2bb21d21-b32f-445b-8d9d-47078092f9c2/coverage.cobertura.xml` ; inventaire `coverage-summary.json` adjacent.
- Résultats globaux : `artifacts/coverage-post-prs/final/global.trx`.
- Organisation miroir : `artifacts/coverage-post-prs/mirror-inventory.json`, produite par `tools/tests/Test-TestLayout.ps1`.
- Catalogue vérifié contre l'assembly reconstruite : **177 outils LLM**.
- [Mesure à 100 % avant la persistance native](test-coverage-pre-native-persistence.md) : historique, distinct de ce passage.

Aucune source du complément n'est exclue ; `ProviderTests.exe` est le seul exécutable de simulation exclu. Le collecteur mesure le processus VSTest, pas le code exécuté dans Excel ou SOLIDWORKS. Les doubles natifs permettent de vérifier les erreurs et frontières du code ; 100 % de couverture ne qualifie pas toutes les commandes dans chaque hôte.

Les [qualifications natives](../reference/native-qualification.md) et le [diagnostic du placement](../chat-persistence-investigation.md) précisent leurs dates et périmètres. Les 24 DesignSurface validées précédemment et les anciens essais SOLIDWORKS restent historiques. L'audit IntelliSense continue avec Luna dans son worktree documentaire.

## Reproduire la mesure

Voir [le projet de tests](../../tests/README.md) pour VSTest, la couverture et les hôtes opt-in. Les rapports détaillés restent locaux sous `artifacts/`, ignoré par Git.
