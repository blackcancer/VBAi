# Documentation VBAi

État vérifié le **28 septembre 2026**, après intégration de `codex/chat-ux` dans `main` (`611dcb5`). Le produit s’appelle VBAi ; la solution, l’assembly, le namespace et les identifiants COM conservent le nom CodexVBE.

## Utiliser le complément

| Document | Contenu |
| --- | --- |
| [État du projet](project.md) | Objectif, fonctionnalités disponibles et validation par hôte |
| [Installation](installation.md) | Compilation, bibliothèque COM, inscription et diagnostic |
| [Conversation](chat-ui.md) | Sessions, contexte, modes, modèles, raisonnement et retour arrière |
| [Fournisseurs](providers.md) | Configuration, authentification, catalogues et limites des transports |
| [GitHub](github-integration.md) | Compte, dépôts, commits, branches, fusions et import VBA |

## Développer et vérifier

| Document | Contenu |
| --- | --- |
| [Architecture](architecture.md) | Projets, dossiers et convention miroir des tests |
| [Concepteurs WinForms](winforms-designer.md) | Vues éditables et séparation entre disposition et données |
| [Localisation](localization.md) | Langues, ressources et maintenance des traductions |
| [Outils LLM](reference/vbe-tools.md) | Catalogue extrait du code : permissions et paramètres requis |
| [Extensions IDE](reference/functional-extensions.md) | Nouveaux contrats et qualifications natives |
| [Concepteur VBE](reference/designer.md) | Propriétés MSForms, conteneurs, listes et récupération |
| [Couverture des tests](test-coverage.md) | Dernière mesure globale et preuves Excel/SOLIDWORKS |
| [Travaux restants](roadmap.md) | Lacunes de tests et de qualification native |
| [Projet de tests](../tests/README.md) | Commandes VSTest et hôtes activés explicitement |

## Lire les états

**Implémenté** signifie présent dans le code. **Validé** précise le scénario et l’hôte testés. **NOT_RUN** signifie non exécuté. Une couverture instrumentée de 100 % ne prouve pas la compatibilité de toutes les fonctions avec tous les hôtes.

Les [archives](archive/README.md) conservent les journaux d’exploration, les décisions de conception et les anciens inventaires. Elles servent de preuve historique ; les anciens chiffres et passages « à faire » ne remplacent pas les guides actuels. Les rapports détaillés sous `artifacts/` restent locaux et ne sont pas versionnés.

## Maintenir la documentation

- Actualiser le guide concerné lorsqu’un comportement change ; conserver les expériences datées dans les archives.
- Mettre à jour `test-coverage.md` depuis le rapport global, avec le commit testé et les limites natives.
- Régénérer ou vérifier le catalogue des outils contre `LlmVbeTools.Definitions` après ajout d’une commande.
- Vérifier les liens relatifs après un déplacement et conserver UTF-8 ainsi que les fins de ligne CRLF.
