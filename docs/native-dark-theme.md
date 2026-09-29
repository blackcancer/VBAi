# Thème sombre natif du VBE

## Récupération des couleurs modifiées — 29 septembre 2026

Le message « Native editor colors were changed after the theme was applied » pouvait bloquer le chargement avec un fichier de récupération valide : les couleurs courantes différaient de la palette enregistrée. Le complément réconcilie désormais chaque couleur de texte, de fond et d’indicateur. Une valeur encore imposée par le thème retrouve sa valeur initiale ; une valeur modifiée depuis est conservée comme nouvelle préférence.

Avant tout changement dans le VBE, le nouvel état remplace atomiquement le fichier de récupération et conserve son contenu précédent dans une archive `palette-{version}.json.previous-{identifiant}` du même dossier. Un échec de remplacement conserve l’ancien fichier. La suppression de la récupération active reste conditionnée à une restauration relue et vérifiée dans Options. Les fichiers invalides et les catégories incompatibles restent refusés.

Qualification : **36 tests ciblés réussis** et **1 parcours Excel natif réussi**, avec application, relecture, archivage et restauration complète des couleurs initiales. Le parcours utilise une récupération temporaire, refuse les hôtes déjà ouverts et ferme son propre classeur sans enregistrer ni exécuter de macro. Preuves : `artifacts/palette-diagnostic/contracts-accepted/palette.trx` et `artifacts/palette-diagnostic/native/native.trx`. Les sections suivantes conservent l’historique des qualifications précédentes.

Le chargement normal de la DLL installée dans un second Excel jetable a également réconcilié le fichier utilisateur `palette-7.01.json`, conservé son ancienne version dans une archive et confirmé « Native editor palette restored and verified ». L’erreur signalée n’est pas reproduite. Preuve : `artifacts/palette-diagnostic/normal-startup.log` ; Excel est ensuite fermé sans sauvegarde.

## Intégration de la PR #8 dans main

Le 28 septembre 2026, l'intégration conserve Monaco issu de la PR #7 et ajoute le thème natif jusqu'au commit `264e432`. L'option reste **expérimentale et désactivée par défaut**. Les captures et essais Office décrits ci-dessous viennent de la branche auteur ; ils n'ont pas été répétés pendant cette intégration, car une autre session utilise Excel.

Contrôles effectués sur main : compilation isolée sans erreur ni avertissement ; **49 tests ciblés réussis**, **32 concepteurs WinForms** et **189 miroirs pour 253 fichiers de production**. Le programme C++ passe **20 cycles** sur des fenêtres synthétiques : périmètre du dessin, barre créée tardivement, refus du mauvais thread, destruction, restauration des imports et équilibre des ressources GDI. Le chargeur vérifie extraction, SHA256, ABI et réutilisation de la DLL **sans activer les hooks Office**. Preuves sous `artifacts/pr8-integration/`.

Correctifs d'intégration : découverte du SDK Windows par son emplacement enregistré (D: sur ce poste), conservation de l'éditeur Monaco actuel, lecture du HWND optionnel protégée dans le démarrage de l'add-in, complétion des quatre nouveaux textes dans les treize catalogues, et arrêt natif confirmé avant de détruire les ressources. Si l'arrêt échoue, palette, HWND, hooks et état actif sont conservés ; la désactivation signale l'échec au lieu d'annoncer une restauration réussie.

Le mode `ForceDark`, global au processus, est réservé à `VBAi_NATIVE_DARK_EXPERIMENT=1`. La case des paramètres applique les traitements ciblés aux fenêtres VBE sans imposer ce mode au processus Office. Le rendu de ce parcours modifié reste à qualifier dans les hôtes réels. Le moteur C++ possède ses propres vérifications et **n'est pas mesuré par Coverlet**, qui mesure le complément C# ; voir le [bilan courant](test-coverage.md).

## Historique de la branche — validation en cours

Le [pilote de rendu direct](native-theme-renderer-pilot.md) ajoute le dessin GDI des onglets Propriétés avec leurs couleurs finales et leur police native ; quatre cas de test ciblés et un passage Excel ont été vérifiés. La trace des appels VBE7 confirme le dessin du code hors WM_PAINT. Le clignotement de toute la barre d'outils est confirmé par capture ; l'optimisation des rafraîchissements qui laissait certains boutons clairs reste désactivée par défaut. Le code et les autres surfaces Office conservent leur moteur de conversion provisoire ; la netteté et la stabilité globales restent à finaliser.

La DLL installée inclut le rendu des Propriétés par ligne, les popups MsoCommandBarPopup et le garde-fou de version Windows. Les sections suivantes conservent l'historique des essais ; leurs résultats ne constituent pas une qualification complète SOLIDWORKS.

Le contrôle de réversibilité a révélé une réentrance des transactions de palette pendant le pompage des messages STA : deux mises à jour pouvaient se disputer le fichier de verrouillage. `VbeNativePalette` sérialise maintenant les mises à jour dans le processus, capture la cible de chaque transaction et diffère une nouvelle demande. Compilation réussie, sans erreur ni avertissement.

La sonde accepte désormais un thème initialement actif, conserve une copie de la récupération et vérifie son identité SHA256 au retour. Elle attend la fin de l'application initiale avant d'ouvrir les paramètres. Le délai précédent de 20 secondes pouvait déclencher une récupération alors que la transaction était encore active ; il est porté à 60 secondes.

`artifacts/vbe-native-settings-awaited-cycle` : application initiale, désactivation, réactivation et parcours des 35 lignes/onglets Propriétés effectués. Le retour suivant au thème clair échoue avec « The native color formatting page was not found ». Le cycle complet n'est donc **pas validé**. Les préférences initiales sont restaurées ; la sauvegarde de palette conservée a le même SHA256 que la copie initiale (`4A99F62FF2D012BEC8FEE2675B1928BB434697CC2027619A828C49F877DF296A`). Aucun arrêt forcé d'Excel n'a été exécuté. L'identification et l'attente de la page de couleurs restent à corriger avant un nouvel essai de réversibilité.

## Objectif

### Correctifs préparés après le signalement des couleurs qui réapparaissent

Le dialogue de palette sélectionne maintenant ses onglets natifs avec TCM_SETCURFOCUS, puis attend la liste 4912 et les trois ComboBox visibles. Les contrôles de couleur peuvent être désactivés selon la catégorie ; cela ne doit pas empêcher leur identification. Le dialogue doit appartenir au VBE, être de classe #32770 et posséder un unique contrôle d'onglets. Le test ciblé `ColorPageSelectionUsesNativeVisibilityAndAllowsDisabledIndicator` réussit (1/1) avec une fenêtre hors écran, une liste cachée portant le même identifiant et un indicateur désactivé. Il ne remplace pas le test du dialogue Options réel. Référence : [TCM_SETCURFOCUS](https://learn.microsoft.com/en-us/windows/win32/controls/tcm-setcurfocus).

SOLIDWORKS PID482784 : pont connecté, DLL chargée depuis bin/Debug/net48, MVID `5cbaa0e0-ab1a-41c6-b745-113ce8526379`. `artifacts/solidworks-native-repaint` montre le titre Projet clair avant RedrawWindow puis sombre après. Les captures écran sont partiellement masquées par l'Explorateur ; cette observation est limitée au titre visible. Aucune macro modifiée ni exécutée. Le passage différé WmRefreshChrome ne traitait que le client et les bordures pour PROJECT/wndclass_pbrs ; il recolore désormais aussi les titres non clients de ces panneaux et de VbaWindow. Compilation isolée réussie sans avertissement/erreur. Ces deux nouveaux correctifs ne sont pas encore chargés dans l'instance SOLIDWORKS et restent à vérifier en situation.

Habiller les fenêtres natives du Visual Basic Editor dans Excel et SOLIDWORKS, en complément du thème déjà appliqué aux fenêtres VBAi. Le réglage est explicite, désactivé par défaut et réversible pendant la session.

Microsoft documente la personnalisation des couleurs du code dans l’onglet **Format de l’éditeur**, mais le VBE n’expose pas de thème sombre complet. Le chrome historique mélange contrôles Win32 standards, surfaces Office personnalisées et fenêtres propres au runtime VBA.

## Première implémentation expérimentale

Le réglage **VBE natif sombre (expérimental)** se trouve dans Paramètres > Apparence. Il est enregistré dans `settings.json` sous `NativeVbeDarkTheme` et s’applique immédiatement.

La classe `VbeNativeTheme` :

- habille les fenêtres VBE ciblées après activation explicite du réglage ; le mode global du processus est réservé à la variable d’expérience ;
- limite l’habillage aux fenêtres appartenant au VBE et ignore les contrôles WinForms/WPF de VBAi ;
- applique le cadre sombre DWM, les thèmes natifs, les fonds et textes des arbres/listes ainsi que les surfaces de contrôles standards ;
- sous-classe uniquement les fenêtres VBE ciblées pour leurs fonds et messages `WM_CTLCOLOR*` ;
- observe la création et l’affichage de nouvelles fenêtres VBE afin d’habiller les volets et dialogues ouverts après l’initialisation ;
- retire les sous-classes, restaure les palettes système et la préférence du processus à la désactivation ou à la déconnexion ;
- ne modifie pas les projets VBA ni les fichiers hôtes ; les couleurs de code sont appliquées via le dialogue Options et leur sauvegarde de récupération est conservée jusqu’à la restauration vérifiée ;
- n’utilise aucun raccourci clavier et ne dépend pas du focus.

Une variable `VBAi_NATIVE_DARK_EXPERIMENT=1` reste disponible pour les sondes jetables sans modifier les préférences utilisateur.

## Résultat observé

Validation du 28 septembre 2026 dans Excel sur Windows 10 22H2, build 19045 :

| Surface | Résultat |
| --- | --- |
| Barre de titre principale | Sombre |
| Espace de travail MDI | Sombre |
| Explorateur de projets | Fond et texte sombres |
| Grille Propriétés | Fond et texte sombres |
| Fenêtre Exécution | Fond non garanti : clair dans le dernier essai avec fenêtre de code rouverte |
| Barres de défilement | Sombres |
| Panneau VBAi | Thème sombre existant conservé |
| Menus et barres d’outils | Restent clairs |
| Bandeaux des volets ancrés | Restent clairs |
| Zone de code | Reste contrôlée par Format de l’éditeur |
| Grilles personnalisées Variables locales/Espions | Cellules encore claires |

Les menus, barres d’outils, bandeaux de volets et certaines grilles sont dessinés directement par l’ancien moteur Office/VBA. Les repeindre intégralement demanderait de remplacer leur dessin natif et augmenterait fortement le risque de régression sur le focus, les commandes, le DPI et les différentes versions de VBE. Cette étape n’est pas incluse dans le premier jalon.

## Compatibilité et risque

`DWMWA_USE_IMMERSIVE_DARK_MODE` est documenté à partir de Windows 11. Le poste de validation utilise Windows 10 ; la barre de titre repose donc aussi sur le comportement de compatibilité de l’attribut historique.

L’activation des contrôles Win32 sombres utilise les points d’entrée privés `SetPreferredAppMode` et `AllowDarkModeForWindow` de `uxtheme.dll`. Microsoft ne garantit pas leur ABI. L’implémentation vérifie leur présence et refuse l’activation si elles sont absentes. Cette dépendance justifie le statut expérimental et l’absence d’activation par défaut.

Avec la variable d’expérience, le mode préféré est global au processus hôte. L’add-in conserve sa valeur précédente et la restaure après un arrêt natif confirmé. Le réglage normal des paramètres ne force pas cette préférence globale. Malgré le filtrage par fenêtre, certains menus contextuels ou dialogues créés par l’hôte peuvent reprendre la préférence globale tant que le thème est actif.

## Vérification reproductible

La sonde crée une instance Excel et un classeur jetables, peut ouvrir une fenêtre de code, capture le VBE avant/après et ferme normalement Excel sans sauvegarder :

```powershell
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File .\tools\probes\Test-VbeNativeDarkChrome.ps1 -InProcessAddInExperiment -OpenCodeWindow
```

Les sorties sont écrites sous `artifacts/vbe-native-dark-chrome/`, ignoré par Git. La preuve visuelle doit accompagner les résultats : un code retour réussi de `SetWindowTheme` ne prouve pas que la surface a réellement changé.

La compilation Debug x64 passe sans avertissement. Les DesignSurface de Chat, Paramètres, Approbation et Git ont été chargées lors du premier jalon.

La sonde accepte aussi `-CycleCodeWindow` avec `-OpenCodeWindow`. Elle ferme et rouvre le volet de code, enregistre `reopened.png` et les hiérarchies natives `windows.json` / `windows-reopened.json`. Le dernier essai a réussi et Excel s’est fermé normalement. Les captures se trouvent sous `artifacts/vbe-native-dark-lifecycle/`.

Le suivi retire maintenant les handles détruits afin de permettre leur réutilisation. La désactivation répétée ne restaure la préférence globale que si une activation l’avait modifiée. Les séparateurs `VBSlider` rejoignent les surfaces dont le fond est pris en charge ; leurs commandes de redimensionnement restent natives.

La cartographie distingue `MsoCommandBar` et `MsoCommandBarDock` (menus et outils Office), `VbaWindow` (code, Exécution et grilles de débogage), `PROJECT` avec `SysTreeView32`, `wndclass_pbrs` (Propriétés) et `VBSlider` (séparateurs). Le même nom de classe `VbaWindow` couvre plusieurs surfaces : une règle de peinture universelle ne suffit pas.

La suite approuvée est décrite dans [le plan de l’éditeur moderne](modern-editor.md).

## Travaux suivants

### Cycle complet depuis les paramètres

La sonde `-SettingsThemeCycle -OpenCodeWindow` utilise le vrai menu `VBAi.Settings`, sélectionne la page Apparence, modifie la case native et enregistre. Elle n'utilise pas la variable d'expérience. Le réglage est relu sur disque et le test attend la vérification du service différé avant de poursuivre.

Le cycle `artifacts/vbe-native-settings-cycle/` a réussi : activation depuis les paramètres, palette appliquée et relue, capture du VBE sombre, désactivation depuis les paramètres, palette restaurée et relue, puis retour visuel au VBE clair. Excel a quitté normalement. Le fichier de paramètres du poste a été sauvegardé à côté de l'original, restauré octet pour octet (empreintes comparées), puis sa copie temporaire a été supprimée. Aucun fichier de récupération de palette ne reste ; seul le fichier de verrou vide peut subsister.

La capture de désactivation contient la fin de l'animation de fermeture d'Options. À la suite de cette observation, le service attend désormais que son dialogue soit invisible avant de déclarer la visite terminée ; la sonde laisse également finir l'animation avant la capture. Cette attente supplémentaire est compilée et sera exercée au prochain cycle groupé.

### Restauration exacte des contrôles natifs

Les couleurs des arborescences, listes et contrôles RichEdit sont mémorisées avant leur remplacement et restituées à la désactivation, au lieu de réimposer les couleurs Windows par défaut. Les valeurs natives qui désignent une couleur système sont conservées. La sauvegarde d'un contrôle n'est pas remplacée par une activation répétée ; elle est supprimée après restauration ou destruction du handle.

Deux tests utilisant de vrais contrôles Windows passent : arborescence avec couleurs personnalisées puis couleur système, et liste avec trois couleurs distinctes (fond, texte, fond du texte). Cela vérifie les messages de restitution ; le parcours complet d'activation/désactivation depuis les paramètres du VBE reste à exercer. Le texte de l'option a été actualisé pour décrire l'application différée et la sauvegarde des couleurs.

### Boutons des listes déroulantes

Le bouton de déroulement est maintenant dessiné indépendamment du texte, à partir du rectangle et de l'état fournis par `GetComboBoxInfo`. Le contrôle natif reste responsable de l'ouverture, de la sélection et des notifications. Les boutons désactivés, survolés ou appuyés ont des couleurs distinctes ; le calcul de taille utilise le rectangle natif.

`artifacts/vbe-native-dark-combos/live.png` montre les flèches sombres des deux listes du code et de celle des Propriétés. La sonde `-CheckNativeCombos` a ouvert puis fermé les trois listes avec `CB_SHOWDROPDOWN`, vérifié leur état ouvert et confirmé que la sélection était conservée. `combo-open-1.png` montre la liste des procédures ouverte. Les réglages d'origine ont été restaurés et relus, puis Excel a quitté normalement. Un second trait de contour a été ajouté après cet essai pour couvrir le relief clair intérieur ; il reste à revoir lors du prochain contrôle visuel groupé.

### Cadres, marge du code et fenêtre Exécution

Les cadres non clients sont désormais dessinés séparément, sur un maximum de deux pixels par bord. La zone client et le texte ne sont pas recopiés pour cette opération. La fenêtre Exécution est identifiée au démarrage avec le type VBE `5`, puis suivie par handle ; elle reçoit le rendu anthracite du code. Les grilles Variables locales et Espions conservent leur traitement distinct.

La marge native claire du code est détectée séparément pour convertir son fond tout en gardant les marqueurs visibles. `artifacts/vbe-native-dark-gutter/debug-break-live.png` montre à l'écran le fond anthracite, la marge sombre, la flèche d'exécution jaune, la ligne `Stop`, la sortie `123` et la variable locale. La macro a repris normalement ; le service a restauré les couleurs et vérifié leur relecture, puis Excel a quitté sans arrêt forcé.

Un rafraîchissement différé a aussi été ajouté après saisie, fin de sélection et défilement, car le VBE peut dessiner directement sans émettre `WM_PAINT`. Le simple survol du code ne recopie plus toute sa surface. Ces chemins de saisie et de défilement restent à qualifier en interaction réelle. Les listes déroulantes et certains boutons conservent encore des surfaces claires.

### Intégration de la palette native au réglage

`VbeNativePalette` est maintenant raccordé au réglage du thème par `VbeNativeTheme` et `AddIn`. Une demande est différée jusqu'à ce que le VBE soit visible et disponible, après fermeture des paramètres. À la déconnexion, le minuteur est arrêté : aucun dialogue n'est ouvert pendant la fermeture de l'hôte.

Le fichier `%LOCALAPPDATA%/VBAi/native-theme/palette-{version VBE}.json` contient les dix catégories d'origine et la palette attendue. Il est créé avant toute modification, avec écriture temporaire vidée sur disque puis déplacement sans remplacement. Une validation du schéma, de la version, des catégories et des indices précède son utilisation. Un fichier verrouillé sérialise les opérations de lecture/modification/restauration.

La désactivation restaure les valeurs d'origine seulement si les valeurs actuelles correspondent à l'origine ou au thème enregistré. Une personnalisation intermédiaire provoque un signalement et conserve la sauvegarde. La suppression du fichier de récupération intervient uniquement après relecture des couleurs restaurées dans un dialogue Options rouvert. Le verrou sérialise les transactions ; le comportement de plusieurs hôtes simultanés partageant leurs préférences VBA reste à qualifier.

Le dialogue est recherché parmi les nouvelles fenêtres appartenant au VBE ciblé. Les catégories sont sélectionnées avec `LB_SETCURSEL` et la notification parent `LBN_SELCHANGE`, car la sélection programmée seule ne produit pas cette notification ([documentation Microsoft](https://learn.microsoft.com/en-us/windows/win32/controls/lbn-selchange)). Le premier essai a montré que `CommandBarControl.Execute` peut revenir avant l'ouverture effective du dialogue : le service attend désormais la fin du worker tout en traitant les messages du thread UI.

**Preuve obtenue :** `artifacts/vbe-native-palette-production3/production-palette-applied.txt` et `production-palette-restored.txt` attestent l'application/restauration et leurs relectures par le service de production dans Excel jetable. Excel a quitté normalement. Huit tests ciblés passent, dont refus des sauvegardes incohérentes, refus d'écraser l'original et préservation des personnalisations intermédiaires. Le cycle complet par la case des paramètres dans l'hôte reste à vérifier ; la sonde appelle directement le même service avec un fichier de récupération sous `artifacts`.

### Correction des surfaces et fond anthracite (28 septembre)

- Les onglets Propriétés conservent leur contrôle natif et utilisent un rendu classique recolorié ; les deux sélections ont été exercées avec `TCM_SETCURFOCUS`, sans raccourci clavier. Le traitement de la liste et des onglets est limité aux enfants de `wndclass_pbrs`, pour ne pas interférer avec les listes du dialogue Options.
- Un rafraîchissement différé après le dessin Office traite le bandeau VBAi et les surfaces natives ; il est coalescé par fenêtre. Le bandeau apparaît sombre dans `vbe-native-dark-surfaces/after.png`.
- Le recoloriage des grilles préserve un fond déjà sombre : le noir de la fenêtre Exécution ne doit pas être inversé en gris clair.
- Le fond de code noir est maintenant remplacé au rendu par l'anthracite `#20242B`, conformément à la demande utilisateur. Le rendu s'appuie encore sur la palette native sombre appliquée par la sonde. Une sélection de couleurs source est traduite vers les valeurs de référence de Visual Studio ; le lissage des caractères reste à améliorer.
- **La capture `PrintWindow` ne prouve pas le rendu final différé.** `artifacts/vbe-native-dark-anthracite-live/live.png`, capturée à l'écran, montre le fond anthracite, alors que les captures synchrones peuvent encore montrer le fond noir. Le contrôle du rendu réel devra couvrir saisie, sélection, défilement et débogage.
- L'essai d'extension aux listes déroulantes, barres de défilement et à tout le cadre a été retiré : il produisait des artefacts. Certaines bordures, flèches et marges restent donc claires.

Compilation réussie sans avertissement. Le dernier essai a restauré les réglages avec relecture (`RestoreVerified: true`) et Excel a quitté normalement. Un essai intermédiaire a échoué lors de la relecture de restauration et a dépassé le délai de fermeture ; Excel a ensuite quitté naturellement, sans arrêt forcé. Le thème reste expérimental et incomplet ; l'intégration durable de la palette native, les contrôles encore clairs et la qualification SOLIDWORKS restent à terminer.

La conversion du code couvre désormais aussi les rampes de gris, vert et bleu-vert de l'anticrénelage, avec des tables calculées une seule fois. Cinq tests limités au thème passent, dont la monotonie des rampes et leur stabilité après plusieurs passages ; les aplats jaune et rouge du débogage sont conservés. Cette vérification mathématique ne valide pas encore les franges ClearType multicolores ni le rendu lors de saisies et défilements réels. Aucun nouvel essai Excel n'a été lancé pour cette modification isolée.

### Avancement de la finalisation native

Un rendu supplémentaire `VbeNativeChrome` intervient après le dessin natif des barres Office, bandeaux et grilles. Il conserve les fenêtres et leur comportement, puis convertit les couleurs du dessin. Cette méthode reste en qualification : certaines icônes deviennent moins colorées et les repeints partiels doivent être contrôlés. Les textes ClearType ont nécessité une conversion de luminance pour éviter des franges colorées.

Les captures `artifacts/vbe-native-dark-chrome-render3/` montrent les menus principaux, les bandeaux Projet/Propriétés/débogage, les onglets Propriétés et les grilles Variables locales/Espions sombres. Le bandeau VBAi et la surface de code restent clairs.

La sonde `-DebugPreview -OpenCodeWindow` exécute une macro jetable par la commande native VBE 186, atteint `Stop`, vérifie le mode arrêt puis reprend et vérifie le retour en conception. L’essai `artifacts/vbe-native-dark-debug/` a réussi avec la variable locale `sampleValue = 123` visible ; Excel s’est fermé normalement. Il révèle cependant que la surface Exécution redevient claire après `Debug.Print`. La dernière correction prend également en compte le DC fourni lors de `WM_PAINT` ; son effet reste à vérifier dans un nouvel essai.

Quatre tests limités au thème passent, notamment la stabilité des couleurs lors de repeints répétés. Cela ne constitue pas une validation des interactions natives ni une qualification SOLIDWORKS. La finalisation reste ouverte : code, bandeau VBAi, mises à jour en débogage, menus ouverts, sélection/survol, désactivation et restauration du rendu.

Le prototype d’éditeur moderne est suspendu ; son entrée de menu a été retirée pendant la finalisation du thème.

### Palette native du code : preuve obtenue

#### Référence visuelle : Visual Studio Community sombre

À la demande de l'utilisateur, la référence est le thème sombre de **Visual Studio Community**, pour l'édition VB.NET. Les couleurs ont été relevées dans l'installation locale de Visual Studio 18, fichier `Common7/IDE/CommonExtensions/Platform/EditorColors.pkgdef`, thème `{1ded0138-47ce-435e-84ef-9ec1f439b749}`. Les entrées `Text Editor Language Service Items` donnent les mots clés `#569CD6`, les commentaires `#57A64A` et les chaînes `#D69D85`. L'entrée `Plain Text` de `Text Editor Text Manager Items` donne un fond `#1E1E1E` et un texte `#DCDCDC`. Ce sont les valeurs du fichier fourni avec cette installation, pas une lecture des personnalisations éventuelles de l'utilisateur.

La sonde utilise une approximation dans les 16 couleurs fixes du VBE :

| Catégorie VBE | Texte | Fond |
| --- | --- | --- |
| Normal, identificateurs, signets | Gris clair `#C0C0C0` | Noir `#000000` |
| Mots clés | Bleu-vert `#008080` | Noir |
| Commentaires | Vert `#008000` | Noir |
| Sélection | Blanc `#FFFFFF` | Bleu foncé `#000080` |
| Erreur de syntaxe | Rouge `#FF0000` | Noir |
| Point d'arrêt | Blanc | Rouge foncé `#800000` |
| Point d'exécution, retour d'appel | Noir | Jaune `#FFFF00` |

Le bleu pur natif est trop sombre sur noir et le cyan vif de la première preuve trop saturé ; le bleu-vert constitue un compromis. Le vert atténué réduit aussi l'effet fluorescent des commentaires. Le contraste reste inférieur à celui des nuances de Visual Studio : cette approximation n'est pas une reproduction exacte de son thème. Le VBE ne propose pas de catégorie indépendante pour les chaînes ou les nombres dans ce dialogue ; ils conservent la couleur native de leur catégorie. Les couleurs d'indicateur et les réglages de police ne sont pas modifiés.

Cette adaptation concerne encore la **sonde**, avant intégration au réglage de production. Elle ne modifie aucun binaire de VBA.

Essai ciblé du 28 septembre : `artifacts/vbe-native-palette-vs-inspired/after.png` confirme les trois couleurs de syntaxe sur le module d'essai. La restauration des dix catégories est relue avec `RestoreVerified: true` et Excel a quitté normalement. Aucun test général n'a été relancé pour cette adaptation. La capture révèle encore une surface Exécution vide claire et le bandeau VBAi clair : la finalisation du thème reste ouverte.

Le 28 septembre, la sonde `-PreviewPalette -DebugPreview -OpenCodeWindow` a appliqué temporairement les couleurs depuis les contrôles réels de l'onglet Format de l'éditeur. Elle utilise UI Automation pour sélectionner les catégories, puis les messages de sélection natifs des ComboBox et leur notification au parent. Les sélecteurs appartiennent à une page imbriquée du dialogue : `GetDlgItem` sur le dialogue principal ne permet pas de les lire correctement.

La capture `artifacts/vbe-native-palette-preview2/debug-break.png` montre le code et Exécution sur fond noir pendant l’arrêt : mots clés cyan, texte clair, ligne courante jaune avec texte noir, et sortie `123` lisible. La macro a repris et est revenue en mode conception. Les valeurs originales des dix catégories sont sauvegardées dans `palette-before.json`, restaurées par les mêmes contrôles puis relues dans un dialogue rouvert. `palette-restore-verified.json` contient `RestoreVerified: true`. Excel a quitté sans arrêt forcé.

Cette palette n’est encore appliquée que par la sonde : elle n’est pas intégrée au réglage de production. L’intégration doit conserver une sauvegarde durable des réglages et permettre leur restauration à la désactivation, sans écraser une personnalisation intermédiaire. La palette fixe du VBE impose ici un fond noir ; aucun fichier VBE.DLL n’a été modifié.

Le test précédent `vbe-native-dark-debug2` confirme que la seule correction du DC de `WM_PAINT` ne suffit pas à conserver Exécution sombre après `Debug.Print`. Le processus Excel de ce test a mis plus de cinq secondes à quitter ; il a ensuite terminé naturellement. La sonde attend désormais jusqu’à trente secondes avant de signaler une fermeture incomplète.

1. Qualifier l’activation/désactivation répétée dans une même instance.
2. Vérifier les fenêtres créées après activation : Options, Explorateur d’objets, Espions, Variables locales, Pile des appels et dialogues modaux.
3. Valider SOLIDWORKS 2020 et comparer son VBE à Excel.
4. Décider séparément de la palette sombre de la zone de code via Format de l’éditeur.
5. Évaluer les menus/barres/bandeaux seulement si une solution préserve commandes, focus, accessibilité, DPI et retour arrière.

## Ajustement du fond de code — 28 septembre 2026

Le fond de code et de la fenêtre Exécution est éclairci de `#20242B` à `#282D35`, dans la même famille anthracite gris bleuté que les panneaux natifs. Les rampes anticrénelées du texte, des commentaires et des mots clés partent de cette nouvelle teinte. Les fonds de sélection, de point d’arrêt et de ligne courante sont conservés. Validation visuelle dans un VBE rechargé encore nécessaire.

## Persistance des couleurs — correction en cours

Le suivi EVENT_OBJECT_SHOW réexamine maintenant les enfants du panneau affiché : leur création peut précéder leur rattachement au VBE, donc leur premier événement ne suffit pas à les reconnaître. Les MsoCommandBar et MsoCommandBarDock participent aussi au rafraîchissement différé après le dessin du parent Office.

La compilation isolée dans `artifacts/theme-persistence-build` réussit sans avertissement ni erreur. Le premier remplacement de la DLL installée a échoué sur un verrou de fichier ; cette correction n'est donc pas encore qualifiée en exécution.

L'essai préalable `artifacts/vbe-native-color-persistence` confirme la restauration de palette et la fermeture normale d'Excel. Les trois listes ComboBox s'ouvrent sans changer la sélection, mais les captures des deux premières sont partiellement occultées : elles ne prouvent pas la conformité visuelle des listes. Le menu Fichier et les dialogues secondaires ne sont pas couverts par cet essai. Le thème reste en cours de finalisation.

### Vérification du correctif chargé

La compilation hors sandbox a remplacé la DLL installée avec succès (0 avertissement, 0 erreur). Restart Manager, interrogé sans demande de fermeture, n'a pas signalé d'utilisateur de la DLL ; cela ne permet pas d'attribuer rétrospectivement le verrou initial à SOLIDWORKS. L'utilisateur a ensuite confirmé sa fermeture, également constatée par la liste des processus.

`artifacts/vbe-native-color-persistence-fixed` : essai terminé avec code 0, trois listes ouvertes sans changement de sélection (`combo-results.json`), deux onglets Propriétés parcourus, code fermé puis rouvert (`reopened.png`). Les captures inspectées montrent le fond de code anthracite et la liste des procédures ouverte. La restauration de palette a été vérifiée dans un dialogue rouvert. Excel a quitté normalement. Les menus Fichier/Édition et les dialogues secondaires ne sont pas couverts par cette validation.

### Dialogues natifs et menus : non qualifiés

Le traitement des boîtes `#32770` appartenant au VBE couvre maintenant le fond, les notifications de couleurs des enfants et le cadre DWM, avec retrait du cadre sombre à la désactivation. Compilation réussie. La sonde `CheckNativeDialogs` n'a pas fourni une capture du dialogue demandé : son image montre une liste du code et ne constitue donc pas une preuve de conformité. Sa restauration initiale a été refusée (éditeur invisible ou modal) ; une session de récupération distincte a ensuite restauré la palette et vérifié sa relecture, puis Excel a quitté. La sauvegarde a été supprimée seulement par le service après vérification. Ne pas déclarer les boîtes Propriétés/Options terminées.

La sonde `CheckCommandMenus` utilise ShowPopup sur un menu existant mais renvoie E_FAIL dans le VBE testé. Elle doit être remplacée ou corrigée avant toute conclusion sur les menus. Le helper de capture doit aussi attendre et identifier précisément la fenêtre attendue, au lieu de se fier à un délai fixe. Ces limitations sont celles des sondes, pas une validation du thème.

### Échec du contrôle visuel interactif

La capture Windows Graphics Capture du VBE a échoué deux fois après récupération de la fenêtre (`FrameArrived timed out`, puis `window capture timed out`). Aucun clic ni saisie n'a été exécuté. La fermeture demandée à la fixture a ensuite reçu RPC_E_CALL_FAILED (0x800706BE) : il ne faut pas qualifier cet arrêt de fermeture normale. Excel n'est plus présent. Le journal Application contient un événement 1000 à 13:22:05 (EXCEL.EXE, exception 0xe0000002 dans KERNELBASE.dll), suivi d'un événement 1001 OFFICE_MODULE_VERSION_MISMATCH à 13:22:16 pour Office 16.0.20326.20158. La causalité avec le thème n'est pas établie. Cette fixture ne modifiait pas la palette native.

La capture de dialogue est renforcée : comparaison des fenêtres visibles avant/après, vérification du PID et de la chaîne de propriétaire VBE, exigence d'un unique nouveau #32770 avec OK/Annuler, délai borné d'apparition, capture de ce seul dialogue et attente de sa fermeture. La tentative ShowPopup est désormais refusée explicitement plutôt que de fournir une image non probante. Le nouveau helper reste à compiler et exécuter ; aucune conformité visuelle supplémentaire n'est revendiquée.

### Contraste du dialogue Options vérifié

`artifacts/vbe-native-dialog-controls/options-dialog.txt` identifie Options, HWND 2363374, classe #32770, propriétaire VBE vérifié. La capture montre un fond sombre, des libellés lisibles, des onglets sombres et les boutons de commande sombres. Le passage se termine avec code 0, annulation du dialogue et fermeture d’Excel. Aucun changement de palette de code. La barre de titre PrintWindow reste claire ; son rendu réel n’est pas validé. La fenêtre Propriétés et les menus supérieurs restent à qualifier. Voir [la recherche sur Rubberduck et les thèmes natifs](native-theme-research.md).

### Menus supérieurs : classe réelle identifiée et corrigée

L’action d’accessibilité `CommandBarPopup.accDoDefaultAction(0)` ouvre le menu existant dans la sonde, sans raccourci ni commande sélectionnée. Cette API Office est réservée à l’usage interne selon la documentation PIA : elle reste strictement dans la sonde. Le produit n’en dépend pas.

La découverte `artifacts/vbe-native-menu-discovery` identifie `MsoCommandBarPopup` (et trois `MsoCommandBarShadow`) avec un propriétaire appartenant au VBE. Le popup était blanc : notre traitement couvrait MsoCommandBar mais pas cette classe. MsoCommandBarPopup participe maintenant au sous-classement, à la peinture de client et au rafraîchissement différé. Les ombres ne sont pas recolorées.

Après compilation sans avertissement/erreur, `artifacts/vbe-native-menu-dark/menu-0-3.png` montre Fichier sombre. `artifacts/vbe-native-menu-reopen` vérifie ensuite Fichier, Édition et Affichage, chacun ouvert/annulé deux fois. Les six captures sont inspectables ; les deux captures de chaque menu ont le même SHA256. Les HWND sont réutilisés par Office et restent pris en charge. Texte actif, sélection et raccourcis sont lisibles ; les éléments désactivés restent distincts. L’essai termine avec code 0 et Excel quitte normalement. Aucun changement de palette native ni commande de menu exécutée.

Limites : preuve dans Excel sur ce poste, principalement PrintWindow ; pas encore de qualification SOLIDWORKS, de tous les sous-menus, du survol prolongé ni du retour au thème clair avec un popup déjà ouvert. Propriétés flottante reste à traiter/qualifier séparément. La sonde libère désormais aussi ses références aux menus avant Quit.

### Texte trouble signalé dans Propriétés — ouvert

L’utilisateur signale un texte trouble, particulièrement dans Propriétés. Le chemin actuel remappe une image après le dessin natif ; son interaction avec les contours ClearType reste suspecte mais la cause précise n’est pas démontrée. Une tentative de préserver toute ListBox dont le fond est déjà sombre a été compilée et contrôlée dans `artifacts/vbe-native-properties-text`. Elle laisse une cellule (Name) claire : le dessin mélange cellules natives claires et fonds déjà sombres. Cette tentative a donc été retirée. Ne pas considérer le flou comme corrigé. La prochaine investigation doit distinguer dessin des cellules, antialiasing et rafraîchissement partiel, plutôt que supprimer globalement la recoloration.

### Rendu Propriétés par ligne native

La sonde `vbe-native-properties-styles` relève le style 0x54810513 de la ListBox de Propriétés : owner-draw fixe, sans LBS_HASSTRINGS. Lire ses données comme des chaînes serait incorrect. Le traitement est déplacé vers WM_DRAWITEM du parent wndclass_pbrs (ODT_LISTBOX) : pour ODA_DRAWENTIRE/ODA_SELECT, SaveDC protège le contexte, les couleurs d’entrée sont rétablies pour le dessin natif, puis seule la ligne rcItem est recolorée avant RestoreDC. Les changements de focus seuls ne repassent pas dans cette transformation. La recoloration globale du client ListBox est retirée.

Compilation sans avertissement/erreur. `artifacts/vbe-native-properties-row-paint` : deux onglets parcourus, capture inspectée avec cellule (Name) et valeur sombres/lisibles, code retour 0, Excel fermé normalement. Ce résultat évite la cellule claire du premier essai de préservation globale. Il ne prouve pas encore la résolution de toutes les situations de texte trouble : grilles riches, sélection, défilement, édition et polices/DPI restent à qualifier. Référence : [WM_DRAWITEM](https://learn.microsoft.com/en-us/windows/win32/controls/wm-drawitem), [DRAWITEMSTRUCT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-drawitemstruct).

### Propriétés : grille riche, sélection et défilement

La sonde `-CheckPropertyRows` crée un UserForm jetable dans le classeur de test, sans sauvegarde. `artifacts/vbe-native-properties-rich/property-row-states.json` relève 35 lignes : sélections 0, 3, 34, puis 0 ; indices supérieurs 0, 3, 7, puis 0. Les deux onglets Propriétés sont également parcourus. Les captures `property-rows-0.png` et `property-rows-2.png` montrent les noms et valeurs lisibles sur fond sombre avant/après défilement. Au retour en haut, `property-rows-3.png` a exactement le même SHA256 que `property-rows-0.png` : 4CF6044D2BA678D090A1E5959E627CEFB118D053CB9DD82D35B74DBC59F78B5D.

L’essai termine avec code 0 et fermeture normale d’Excel. Il apporte une preuve de stabilité du dessin sur ce parcours, pas une validation de la saisie réelle, de tous les états de focus ni de SOLIDWORKS. Les échantillons de couleur intégrés aux cellules méritent aussi un contrôle de fidélité séparé : la transformation de ligne ne doit pas fausser leur signification.

### Compatibilité des appels privés et validation restant à faire

Avant de résoudre l’ordinal135 de uxtheme, le produit lit la version native Windows et refuse l’ABI antérieure à Windows10 1903 (build18362). Sous1809, cet ordinal correspond à AllowDarkModeForApp(bool), et non SetPreferredAppMode(enum). Un changement futur de version majeure doit être qualifié avant usage. La présence des exports est toujours vérifiée ; ce garde-fou n’apporte pas de garantie Microsoft sur les API privées.

Compilation des projets et test ciblé `NativeWindowsVersionIsReadBeforeUsingVersionDependentOrdinals` réussis (1/1) dans `artifacts/native-theme-version-check`. La DLL installée n’a pas été remplacée pendant l’utilisation de SOLIDWORKS PID475276. La première invocation du filtre, avant ajout effectif du test, ne trouvait aucun test : seule la seconde invocation, avec 1 test exécuté, constitue la preuve.

La campagne de désactivation a été reportée : NativeVbeDarkTheme=true et une sauvegarde palette-7.01.json existent pendant que SOLIDWORKS est ouvert. La sonde SettingsThemeCycle refuse désormais ce contexte avant toute mutation. Une lecture native confirme l’absence de débogueur attaché à ce PID. Les outils MCP Visual Studio requis par le skill solidworks-debug-pack-and-go sont absents de la session : aucune campagne SOLIDWORKS conforme à ce protocole n’a été menée. Les préférences et la sauvegarde existantes sont conservées.
