# Modifier les interfaces dans Visual Studio

Ouvrir `CodexVBE.sln`, compiler la solution, puis ouvrir le fichier `.cs` voulu avec
**Afficher le concepteur** (`Maj+F7`). Ouvrir le `.cs` principal, pas le
`.Designer.cs`. Les fenêtres et sous-vues sont des composants WinForms ; les
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
initiale se règle dans le Designer ; changer de fournisseur ajuste la hauteur au
contenu des vues, tandis que les réponses d’authentification ne redimensionnent plus
la fenêtre. Les champs incompatibles avec le fournisseur sélectionné restent
masqués à l’exécution, dans des lignes AutoSize définies par le Designer.

`ChatWindow`, `VbeApprovalDialog`, `ChatToolWindow` et `CodeDiffView` conservent
également un fichier Designer et un fichier de ressources. Les quatre colonnes
du diff sont créées dans le Designer ; le mode unifié adapte leur présentation.

## Fenêtre À propos

`src/CodexVBE/Ui/AboutWindow.cs` est un formulaire indépendant. Son en-tête,
logo, descriptif, tableau de métadonnées, liens, état et boutons sont déclarés
explicitement dans `AboutWindow.Designer.cs`. Le logo est une ressource image
WinForms éditable dans `AboutWindow.resx`. La version, l’hôte et la langue sont
renseignés à l’exécution sans démarrer le chat ni charger de paramètres fournisseur.

La fenêtre est accessible depuis **Outils → À propos de VBAi** et le menu du chat.
Voir [À propos et support](about.md) pour les informations copiées et les liens.

## Fenêtre de rapport de problème

`src/CodexVBE/Ui/CrashReportWindow.cs` possède son `.Designer.cs` et son `.resx`.
Ses 20 contrôles fixes (description, aperçu, destination, état, progression et boutons)
sont éditables dans le concepteur. Le constructeur Designer ne charge aucun compte
et n’appelle aucun transport. Voir [Rapports de problème](crash-report.md).

## Éléments réutilisables du chat

Les vues suivantes sont dans `src/CodexVBE/Llm/Controls/`. Ouvrir chacune avec
**Afficher le concepteur** pour modifier ses contrôles internes.

| Vue | Contenu éditable |
| --- | --- |
| `ChatInputView` | Emplacement de saisie, aperçu WinForms, police, marges du texte et activation du correcteur |
| `ChatContextChipView` | Puce de référence ou pièce jointe, boutons ouvrir / retirer et tooltips |
| `ChatContextPreviewView` | Groupe et champ de lecture du contexte joint |

Le moteur WPF de saisie est créé à la première utilisation à l’exécution pour
conserver le correcteur orthographique et les interactions clavier. Le constructeur
Designer montre un champ WinForms et ne démarre pas ce moteur. `Font`, `ForeColor`,
`InputPadding` et `SpellCheckEnabled` sont transmis au moteur.

Le diff inline réutilise `src/CodexVBE/Ui/CodeDiffView.cs` : grille, colonnes,
recherche, navigation et options sont construits dans son Designer. `ChatDiffView`
est seulement un adaptateur `WindowsFormsHost` pour le transcript WPF. Le mode
unifié se règle via `UnifiedDiff`. Les hôtes de diff sont libérés lorsqu’une carte
est retirée, virtualisée ou lorsque la fenêtre est fermée.

`ChatContentHost` utilise le concepteur WinForms standard pour l’emplacement du
contenu WPF. Le Designer n’a pas besoin de créer un transcript ou une session
pour éditer cet emplacement.

## Répartition entre Designer et exécution

- Disposition, contrôles fixes, ancrage, tailles, marges et tooltips : `.Designer.cs`.
- Chargement des données, activation des boutons, sélection et opérations : `.cs`.
- Messages du chat, Markdown, actions propres aux messages et suggestions au curseur :
  contenu dynamique WPF dans les emplacements WinForms du `ChatWindow`.
- Puces et aperçus : instances des vues Designer selon les données ; leurs contrôles
  internes ne sont pas reconstruits par le contrôleur.
- Lignes du diff : données virtuelles de la grille Designer partagée par Git et le chat.
- Saisie : moteur WPF interopérable pour le correcteur, dans un hôte fixe Designer.
- Thème et traduction : appliqués à l’exécution ; le Designer conserve ses libellés
  anglais éditables.
- Les constructeurs sans paramètre des vues ne se connectent ni au VBE ni à GitHub.
- `InitializeComponent` déclare et configure explicitement les composants, sans
  boucle ni fabrique de disposition. Les noms, tailles, positions et TabIndex sont
  présents ; les conteneurs suspendent/reprennent leur disposition, et les grilles
  et SplitContainer respectent `ISupportInitialize`.

Modifier une sous-vue en ouvrant directement son fichier : la fenêtre parente ne
sert pas à éditer les contrôles internes d’un UserControl. Ne pas supprimer les
contrôles nommés utilisés par le contrôleur sans adapter leurs références.

## Vérification

`tools/tests/Test-WinFormsDesigners.ps1 -AssemblyPath <chemin de CodexVBE.dll>`
valide **28 surfaces et 282 contrôles enfants** avec le moteur
`System.ComponentModel.Design.DesignSurface` : chargement, redimensionnement,
édition d’une propriété puis sérialisation et rechargement avec
`CodeDomComponentSerializationService`.

Le test place les contrôles déclarés par le Designer sur une racine WinForms
éditable. Charger directement la classe compilée modélise un formulaire hérité
et verrouille ses champs privés ; ce second cas ne prouve pas l’édition du source.
`Test-ChatDesigner.ps1` vérifie en plus les hiérarchies, les constructeurs inertes
et la sélection des contrôles. Ces contrôles locaux ne constituent pas un essai
manuel d’ouverture puis d’enregistrement des 28 sources dans Visual Studio.

`Render-ChatUx.ps1` capture la fenêtre affichée sur le second écran disponible.
La copie écran inclut les HWND des diff natifs, absents d’un rendu bitmap WPF.
Le paramètre `-Theme Light` ou `-Theme Dark` ne change le thème que dans le
processus de démonstration ; la préférence enregistrée reste intacte.

Preuves locales de cette refonte :
`artifacts/designer-refactor/designers/designers.json`, captures sous `screens/`
et résultats unitaires sous `tests/`. Compilation solution : **0 erreur,
0 avertissement**. Suite unitaire complète : **1 115 réussites** ; les **73 tests
chat/diff** ont été rejoués avec succès après les derniers ajustements de thème
et de métriques de police. Aucune macro n’est exécutée par ces essais. La couverture
instrumentée n’a pas été remesurée pour cette refonte.

## Mise à jour ultérieure

L’extension du thème au VBE natif est différée : palette de l’éditeur, sauvegarde
et restauration des réglages, puis étude des fenêtres natives sous Excel et
SOLIDWORKS. Cette refonte ne modifie pas le thème du VBE.
