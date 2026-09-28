# VBAi : objectif et état du projet

État du code **`2197c43`**, vérifié le **28 septembre 2026**.

## Objectif

Intégrer un assistant LLM au Visual Basic Editor pour lire, modifier et déboguer le projet VBA vivant, dans Excel comme dans SOLIDWORKS. Codex est le fournisseur par défaut et utilise l’authentification ChatGPT du CLI ; une clé OpenAI API n’est pas nécessaire pour ce mode. Les autres transports sont décrits dans [Fournisseurs](providers.md).

Le complément cible **VBE 64 bits, .NET Framework 4.8 et Windows**. Le chat, les outils et la session VBE s’exécutent dans le complément COM ; il ne s’agit pas d’un serveur MCP. Le pont local `CodexVBE.<PID>` permet aussi les diagnostics et les tests externes.

## Fonctionnalités présentes

| Surface | Fonctionnalités implémentées |
| --- | --- |
| Projets et références | Inventaire, propriétés, composants, références COM, bibliothèques de types, certificats et commande de signature |
| Modules et classes | Création, import/export, renommage de composants, lecture/édition avec SHA, procédures, événements, recherche, symboles, navigation et historique |
| UserForms | Arbre des conteneurs, contrôles, propriétés typées, police/images, disposition, sélection, duplication bornée, listes, événements et récupération après coupe |
| Exécution et débogage | Compilation, lancement, breakpoint natif, pas à pas, reprise/reset, fenêtres de diagnostic, Exécution et espions |
| Fenêtres de l’éditeur | Volets de code, vues, fenêtres, disposition, barres d’outils, compléments et Explorateur d’objets |
| Conversation | Contexte VBE dynamique, références `#`/`@`, modes Discussion/Plan/Agent, sessions SQLite, choix persistants, résumé de réflexion et rollback |
| Git et GitHub | Compte indépendant du fournisseur IA, export versionné, commits, push/fetch/pull, checkpoints, branches, fusions et PR |

Le [catalogue des 166 outils LLM](reference/vbe-tools.md) fournit les noms et paramètres requis. Les fonctions du pont et les outils du modèle ont des périmètres distincts.

## Gardes et limites

- Les actions identifient le projet et le composant, vérifient le mode et les révisions applicables, puis relisent le résultat.
- Les niveaux d’approbation sont Lecture seule, Demander à chaque action et Automatique. Discussion/Plan bloquent les actions d’édition et d’exécution indépendamment de cette politique.
- Les raccourcis VBE et `SendKeys` sont proscrits. Les commandes utilisent VBIDE, COM et les interfaces natives ; certaines opérations natives reposent sur une géométrie locale vérifiée, dont la qualification DPI/écrans reste limitée.
- Une sauvegarde Excel est qualifiée séparément d’une modification en mémoire. Une modification réussie dans SOLIDWORKS ne prouve pas sa persistance dans le `.swp`.
- Les propriétés COM recensées, les contrôles tiers et les retours asynchrones ne constituent pas une couverture native universelle.

## Dernière validation

| Vérification | Résultat |
| --- | --- |
| Suite globale VSTest avec hôtes activés | **910 réussis, 0 échec, 0 ignoré** |
| Couverture des lignes / branches | **92,94 % / 87,79 %** |
| Compilation | **0 erreur, 0 avertissement** |
| Concepteurs WinForms | **24 chargements et modifications de taille réussis** |
| Excel | Chargement, pont, inspection et sauvegarde/relecture d’un classeur macro jetable |
| SOLIDWORKS 2019 SP5 | Chargement, compilation/exécution, breakpoint, pas à pas et reprise sur `test.swp` ; inventaire et SHA initiaux restaurés |

La capture des variables locales de SOLIDWORKS expose un panneau accessible mais aucune ligne dans ce dernier essai. La lecture de leurs valeurs n’est donc pas validée. Le dernier passage global ne qualifie pas chaque commande dans les deux hôtes.

Voir [les preuves et commandes de mesure](test-coverage.md) et [les travaux restants](roadmap.md). L’[ancien état initial du projet](archive/exploration/project.md) est conservé comme historique.
