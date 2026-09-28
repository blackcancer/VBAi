# Modifier les interfaces dans Visual Studio

Ouvrir `CodexVBE.sln`, compiler la solution, puis ouvrir le fichier `.cs` voulu avec
**Afficher le concepteur** (`Maj+F7`). Ouvrir le `.cs` principal, pas le
`.Designer.cs`. Les vues ci-dessous sont des contrôles utilisateur WinForms ; les
onglets restent regroupés dans leurs fenêtres habituelles.

## Git

Les fichiers sont dans `src/CodexVBE/Git/Views/`.

| Vue | Contenu éditable |
| --- | --- |
| `GitConnectionView` | Lien du document avec un dépôt et une branche |
| `GitChangesView` | Message de commit, sélection des modules, diff et actions de revue |
| `GitHistoryView` | Liste des commits, détails et comparaison |
| `GitBranchesView` | Création, suivi, changement de branche et fusion |
| `GitCheckpointsView` | Création et restauration des checkpoints |
| `GitConflictsView` | Ancêtre commun, versions, résolution et actions de fusion |
| `GitImportView` | Récapitulatif de l’import |
| `GitHubRepositoriesView` | Recherche, sélection et création de dépôts |
| `GitHubPullRequestsView` | Liste des PR et navigation dans leurs détails |
| `GitHubPullComposeView` | Formulaire de création d’une PR |
| `GitHubPullDetailsView` | Description et état de la PR |
| `GitHubPullFilesView` | Fichiers de la PR |
| `GitHubPullCommentsView` | Commentaires et contenu sélectionné |
| `GitHubPullChecksView` | Résultats des contrôles |

`GitWindow` contient le bandeau commun, les onglets et la zone d’état.
`GitHubPane` contient les onglets Dépôts / Pull requests et leur état réseau.
Les fichiers `.Views.cs` relient les événements et les contrôles aux opérations
existantes ; ils ne construisent ni ne déplacent les contrôles.

La connexion dispose désormais de son propre onglet. Le message et le bouton de
commit appartiennent à Modifications Git. Changer d’onglet ne réécrit plus les
hauteurs des lignes de la fenêtre principale.

## Paramètres et autres fenêtres

Les vues de `src/CodexVBE/Llm/Settings/Views/` sont :

- `ProviderSettingsView` : paramètres des fournisseurs et permissions d’édition ;
- `GitHubAccountSettingsView` : compte GitHub des dépôts ;
- `AppearanceSettingsView` : thème de VBAi.

`LlmSettingsWindow` contient ces trois onglets et Enregistrer / Annuler. Sa taille
se règle dans le Designer ; les réponses d’authentification ne redimensionnent plus
la fenêtre. Les champs incompatibles avec le fournisseur sélectionné restent
masqués à l’exécution, dans des lignes AutoSize définies par le Designer.

`ChatWindow`, `VbeApprovalDialog`, `ChatToolWindow` et `CodeDiffView` conservent
également un fichier Designer et un fichier de ressources. Les quatre colonnes
du diff sont créées dans le Designer ; le mode unifié adapte leur présentation.

## Répartition entre Designer et exécution

- Disposition, contrôles fixes, ancrage, tailles, marges et tooltips : `.Designer.cs`.
- Chargement des données, activation des boutons, sélection et opérations : `.cs`.
- Messages du chat, Markdown, cartes de diff et suggestions : contenu dynamique WPF
  dans les emplacements WinForms du `ChatWindow`.
- Thème et traduction : appliqués à l’exécution ; le Designer conserve ses libellés
  anglais éditables.
- Les constructeurs des vues ne se connectent ni au VBE ni à GitHub.

Modifier une sous-vue en ouvrant directement son fichier : la fenêtre parente ne
sert pas à éditer les contrôles internes d’un UserControl. Ne pas supprimer les
contrôles nommés utilisés par le contrôleur sans adapter leurs références.

## Vérification

`tools/tests/Test-WinFormsDesigners.ps1 -AssemblyPath <chemin de CodexVBE.dll>`
charge les 24 surfaces avec le moteur `System.ComponentModel.Design.DesignSurface`
et vérifie l’édition de leur taille. `Test-ChatDesigner.ps1` vérifie aussi les
hiérarchies et la sélection de contrôles. Ce contrôle automatisé ne constitue pas
un essai manuel d’enregistrement de chaque Designer dans Visual Studio.

## Mise à jour ultérieure

L’extension du thème au VBE natif est différée : palette de l’éditeur, sauvegarde
et restauration des réglages, puis étude des fenêtres natives sous Excel et
SOLIDWORKS. Cette refonte ne modifie pas le thème du VBE.
