# Interface compacte et style commun

Le **chat est la référence visuelle** pour les autres fenêtres. Les réglages de
présentation sont déclarés dans les fichiers `.Designer.cs` : aucune fabrique de
contrôles fixes n'est ajoutée dans les contrôleurs.

## Comportement

- Le titre de la conversation, le document et l'historique occupent l'en-tête.
- Le modèle et son effort sont résumés près de la saisie. Un clic révèle les
  sélecteurs existants de fournisseur, modèle et effort, sans changer leur valeur.
- La politique des **éditions VBA** est accessible dans le chat : Automatique,
  Demander à chaque action, Lecture seule. Elle modifie le réglage réellement
  consommé par les outils, est conservée dans les paramètres et est verrouillée
  pendant une intervention. Elle ne remplace pas les politiques propres au
  fournisseur, au sandbox ou aux autorisations d'accès interprojets.
- Les activités en cours s'ouvrent dans le fil ; les résumés de réflexion transmis
  par le fournisseur se développent pendant leur réception. Un repli manuel est
  respecté lors des mises à jour. À la fin, le groupe se replie sauf s'il a été
  explicitement ouvert. Aucun raisonnement privé ou texte fictif n'est ajouté.
- Les commandes fréquentes deviennent des icônes avec infobulles et nom accessible.
  Les décisions ambiguës ou sensibles conservent leur libellé (approbation,
  installation, choix d'une version en conflit). Envoi, arrêt et reprise gardent
  leurs états fonctionnels et leur texte accessible.
- Le menu Monaco regroupe les commandes de débogage et d'assistance sans répéter
  « VBAi » devant chaque entrée. Les identifiants et raccordements restent stables.

## Git et éditeur

`GitWindow` reprend l'en-tête compact du chat, avec les commandes de synchronisation
à droite, les onglets en pastilles de 36 DIP (sélection encadrée, survol, focus clavier) et des cartes pour la connexion et le message de
commit. Les vues restent des UserControls indépendants éditables dans Visual Studio.
Les listes éditables désactivées suivent aussi la palette sombre.

La préférence de thème est commune à Monaco et aux onglets WinForms. Les notifications
Windows sont ramenées sur le thread UI ; Monaco reçoit la même couleur de surface.
Les mêmes onglets sont utilisés dans Git, ses sous-vues, les paramètres et Monaco ;
les modules gardent une croix de fermeture par onglet. Le diff conserve le thème courant, et le contraste élevé est piloté par l'hôte.
`Test-EditorAppearance.ps1` vérifie les couleurs calculées dans le vrai WebView2 en
clair, sombre, puis clair, pour le code, le diff et le retour au code (9 cas).

## Contrôles communs éditables dans Visual Studio

| Élément | Contrôle | Règle |
| --- | --- | --- |
| Commande | `UiActionButton` | Propriétés Designer `Symbol`, `IconOnly`, `Primary` ; clavier, focus, infobulle et accessibilité natifs |
| Sélecteur | `UiComboBox` | Style du sélecteur du chat : rayon de 7 DIP, bordure discrète, chevron, hauteur de ligne de 22 DIP minimum |
| Texte | `UiTextBox`, `UiRichTextBox` | Même bordure arrondie ; édition, sélection, mot de passe et défilement natifs |
| Liste | `UiListBox`, `UiCheckedListBox` | Même contour ; sélection et cases natives |
| Grille | `UiDataGridView` | Même contour, couleurs de diff conservées |
| Carte de saisie | `ChatComposerPanel` | Arrondi de 12 DIP conservé ; `ShowBorder` permet les messages sans cadre |

`ChatChoiceBox` et `ThemedComboBox` héritent du même sélecteur. Les champs, listes
et grilles de Git, des paramètres, des mises à jour et des boîtes de dialogue
emploient ces contrôles dans leurs Designers. La palette est centralisée dans
`UiTheme` : fond, surface, bordure, focus, texte et menus contextuels. Les couleurs
de contraste élevé suivent Windows. Dans le concepteur Visual Studio, la palette
est celle du formulaire et de ses parents : une préférence VBAi sombre ne peint
plus des onglets sombres au milieu d'un Designer clair. La palette de peinture suit les couleurs effectives du parent, même avant que
Visual Studio fournisse un `Site` au contrôle. La détection du mode Designer
empêche également le thème runtime de remplacer les couleurs du formulaire. Les coins sont peints autour des contrôles
natifs ; leur saisie, leur liste déroulante et leurs barres de défilement restent
celles de WinForms.

Les 25 symboles originaux se trouvent dans `assets/icons/commands/`. Ils sont
incorporés dans l'add-in et l'Updater, dessinés à la résolution de l'écran avec la
couleur courante, sans chargement réseau. Le lecteur accepte les chemins absolus
SVG `M`, `L`, `C`, `Z` utilisés par ces fichiers ; il n'est pas un moteur SVG général.

Les gestionnaires des nouveaux événements du chat sont dans `ChatWindow.cs`,
ce qui permet au concepteur de résoudre `ApprovalPicker_SelectedIndexChanged`.

## Sources de conception

- Sélection du modèle et de l'effort à proximité de la saisie :
  [aide Claude](https://support.claude.com/en/articles/8664678-change-the-model-effort-and-thinking-settings).
- Relecture des changements et suivi de l'intervention dans un même parcours :
  [présentation de Codex](https://openai.com/index/introducing-the-codex-app/).
- Organisation de l'interface de contrôle :
  [documentation OpenClaw](https://docs.openclaw.ai/web/control-ui).
- Infobulles sur les commandes représentées par une icône :
  [recommandations Microsoft](https://learn.microsoft.com/en-us/windows/apps/develop/ui/controls/tooltips).

Ces références guident les interactions ; le style conservé est celui du chat VBAi.

## Validation

### Intégration dans main — 29 septembre 2026

La PR #12 est intégrée avec le correctif de récupération de palette native. Le lot corrigé compte **26 tests réussis, 0 échec**, dont les parcours réels WebView2 ; les **57 scénarios JavaScript** d’IntelliSense et d’édition restent verts. Les contrôles de structure passent : **46 concepteurs WinForms**, **27 éléments de métadonnées**, **239 miroirs pour 297 sources**. La documentation XML couvre **5 352 / 5 352 déclarations**, sans erreur syntaxique ; ses compléments ne modifient pas le code exécutable.

Les neuf cas du renderer vérifient les thèmes clair, sombre puis clair dans le code, le diff et le retour au code. Le chargement normal de l’ensemble fusionné dans un Excel jetable confirme la récupération de palette utilisateur sans l’erreur signalée. Il ne constitue pas une qualification visuelle complète du chat ou de SOLIDWORKS. Preuves : `artifacts/pr12-integration/accepted/accepted.trx`, `designers/designers.json`, `metadata/project-items.json`, `editor-appearance/theme-validation.txt` et `artifacts/palette-diagnostic/normal-startup.log`. La mesure globale courante est détaillée dans [le bilan des tests](test-coverage.md).

### Historique de la branche

Les vérifications couvrent la compilation de l'add-in et de l'Updater, les états
du chat, les icônes à 96 et 192 DPI, la stabilité des tailles, les champs natifs
et les Designers. Les captures utilisent des données de démonstration locales ;
elles ne prouvent pas un parcours réel dans Excel ou SOLIDWORKS ni un échange avec
un fournisseur authentifié. La couverture globale de la nouvelle version n'a pas
été recalculée lors de cette évolution visuelle.

Commandes reproductibles : `Render-ChatUx.ps1`, `Render-CompactUi.ps1`,
`Test-WinFormsDesigners.ps1`, `Test-WinFormsProjectMetadata.ps1` et
`Test-VisualStudioWinForms.ps1` dans `tools/tests/`.

Validation complémentaire du thème dans le Designer : 17 tests ciblés réussis,
46 surfaces WinForms chargées et sérialisées, puis contrôle visuel dans Visual
Studio de GitWindow et ModernEditorWindow après rechargement de leurs documents
enregistrés. Les onglets restent clairs sur le formulaire clair, même lorsque
la préférence de thème de VBAi est sombre.

Les actions principales utilisent un fond gris perle, un contenu foncé et une bordure discrète en thème clair, et blanc cassé
en thème sombre, avec des nuances distinctes au survol et à la pression. Cette
palette suit la surface parente dans le Designer et conserve les couleurs système
en contraste élevé. Elle remplace le bleu saturé des boutons sans changer les
indications de focus clavier ni les couleurs sémantiques du contenu.
