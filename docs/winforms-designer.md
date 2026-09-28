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
recherche, navigation et options sont construits dans son Designer. `ChatChangeCardView` contient ce contrôle dans son Designer. `ChatDesignerHost`
adapte les vues natives au transcript virtualisé. Le mode
unifié se règle via `UnifiedDiff`. Les hôtes de diff sont libérés lorsqu’une carte
est retirée, virtualisée ou lorsque la fenêtre est fermée.

`ChatContentHost` utilise le concepteur WinForms standard pour l’emplacement du
contenu WPF. Le Designer n’a pas besoin de créer un transcript ou une session
pour éditer cet emplacement.

## Répartition entre Designer et exécution

- Disposition, contrôles fixes, ancrage, tailles, marges et tooltips : `.Designer.cs`.
- Chargement des données, activation des boutons, sélection et opérations : `.cs`.
- Messages, activités, pièces jointes, actions et suggestions : instances de vues
  WinForms Designer selon les données. Les contrôles internes sont fixes.
- Markdown : contenu et styles du RichTextBox natif défini dans le Designer.
- Transcript : infrastructure WPF virtualisée dans un emplacement fixe ; elle
  héberge les vues WinForms sans construire leurs dispositions.
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

## Mises à jour

`Updates/UpdateWindow.cs` configure les mises à jour et affiche les notes de release.
`src/CodexVBE/Updates/UpdateProgressWindow.cs` est partagée avec le programme de mise à jour
différée. Les deux formulaires possèdent leurs Designer et ressources. Le programme
externe partage les sources et les catalogues, sans charger l’assembly COM. Les
contrôles secondaires et principaux ont des zones distinctes ; la progression reste
dans une ligne dédiée. Voir [Mises à jour](updates.md) pour le protocole installeur.

## Cartes et suggestions du chat

Les vues de `src/CodexVBE/Llm/Controls/Transcript/` sont réellement utilisées
par `ChatWindow.Transcript.cs`, `.Activities.cs`, `.FormRecovery.cs` et `.Composer.cs`.
Ouvrir leur fichier `.cs` avec **Afficher le concepteur** (`Maj+F7`).

| Vue | Contrôles fixes dans le Designer |
| --- | --- |
| `ChatMessageView` | En-tête, copier, créer une branche, message, mémoire, listes de contexte, annulation et correction |
| `ChatActivityGroupView` | Section de chronologie des activités |
| `ChatActivityStepView` | État et détail dépliable d'une étape |
| `ChatChangeCardView` | Module, compte des changements, `CodeDiffView`, boutons de rollback et état |
| `ChatAttachmentView` | Section, contenu joint et navigation vers le code |
| `ChatFormRecoveryView` | Titre, bilan et récupération d'un formulaire |
| `ChatWelcomeView` | Titre, indication et trois suggestions |
| `ChatSuggestionsView` | Liste d'autocomplétion et état |
| `ChatTextContentView` | RichTextBox en lecture seule et menu de copie |
| `ChatLinkView` | Bouton de navigation vers une référence |
| `ChatDisclosureView` | Bouton de dépliage et conteneur de contenu |

`ChatQueuedMessageView`, dans `Llm/Chat/`, définit les actions Envoyer maintenant,
Modifier et Supprimer des messages en attente.

`ChatDisclosureDesigner` expose `ContentPanel` au concepteur : on peut y déposer
un contrôle dans une vue parente. Le diff et les contenus des pièces jointes sont
ainsi déclarés dans leurs Designers, pas ajoutés par le contrôleur à l'exécution.
Les listes variables (messages, références, pièces jointes, étapes et blocs du menu
de rollback) créent uniquement les instances nécessaires aux données.

La saisie conserve son moteur WPF pour la correction orthographique ; sa surface
WinForms est éditable dans `ChatInputView`. La virtualisation du fil et le placement
du popup au curseur restent des mécanismes d'exécution.

Validation locale : chargement des douze vues dans `DesignSurface`, édition du
conteneur imbriqué, affichage d'un message, d'une référence et d'un diff dans
`WindowsFormsHost`, captures natives claires et sombres. Ces tests ne remplacent
pas une ouverture manuelle du projet dans le Designer de Visual Studio.

## Éditeur Monaco, mises à jour et fichiers partiels

`ModernEditorWindow.cs`, `UpdateWindow.cs` et `UpdateProgressWindow.cs` sont les
entrées du Designer dans le projet principal. La progression est partagée avec
le programme `VBAi.Updater`, mais sa surface s'édite uniquement dans
`src/CodexVBE/Updates/`. Les liens de code dans l'Updater sont classés Code et la
ressource est incorporée à la compilation sans ajouter une deuxième surface
Designer : l'ouverture de cette copie dans l'Updater échouait à charger sa
ressource dans Visual Studio.

Les fichiers `ModernEditorWindow.Debug.cs`, `.Language.cs`, `.Save.cs`, `.Tools.cs`,
`GitWindow.Review.cs`, `.Views.cs` et `LlmSettingsWindow.Views.cs` contiennent la
logique et les liaisons aux vues. Ils sont classés **Code** et regroupés sous le
fichier principal ; ce ne sont pas des surfaces Designer supplémentaires.
Les dispositions Git et paramètres restent éditables dans leurs UserControls.

Les métadonnées de mise à jour doivent suivre l'inclusion des fichiers C#.
Les règles de classement des fichiers partiels sont appliquées après les imports
SDK : Visual Studio peut conserver un `SubType=Form` dans le fichier utilisateur
`.csproj.user`, qui prenait auparavant priorité sur les réglages du projet.
L'import explicite de `Sdk.props` / `Sdk.targets` conserve le projet SDK et permet
ces règles finales, sans supprimer les préférences locales.

Les déclarations du Designer emploient des affectations séparées et des délégués
explicites. Les aperçus Designer ne démarrent ni WebView2 ni le parcours de mise à
jour. Le moteur Monaco et ses documents restent initialisés à l'exécution.

Contrôles reproductibles :

- `tools/tests/Test-WinFormsProjectMetadata.ps1` vérifie les éléments évalués par
  MSBuild, y compris le classement Code de la copie partagée dans `VBAi.Updater` ;
- `WindowDesignerCompatibilityTests` vérifie les composants éditables et leur
  sérialisation pour les cinq fenêtres, ainsi que l'absence d'initialisation
  runtime dans les aperçus ;
- `tools/tests/Test-VisualStudioWinForms.ps1 -VisualStudioProcessId <PID>` ouvre
  les fenêtres dans le Designer d'une instance VS associée à ce checkout et
  vérifie leurs racines chargées. Le contrôle ne modifie pas les
  composants et conserve les documents ouverts pour inspection.
