# Couverture VBE restante

Inventaire au 27 septembre 2026 sur `main`. « Présent » désigne une commande du complément, « prouvé » un essai dans le VBE Excel visible. La présence d'un descripteur COM modifiable ne prouve pas que son écriture est sûre ni durable.

Le périmètre inclut les onze familles de [menus intégrés du VBE](https://learn.microsoft.com/en-us/office/vba/language/reference/menus-commands), les menus contextuels, les fenêtres et l'[objet VBE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model) (projets, composants, code, références, compléments, événements et barres de commandes). Les menus documentés par Microsoft varient selon l'hôte et la version ; ce tableau relève les fonctions, sans assimiler chaque action visuelle à une commande LLM nécessaire.

## Ajouts du worktree `codex/chat-ux` — 28 septembre 2026

Le tableau historique ci-dessous décrit la référence `main` du 27 septembre. Les ajouts suivants sont implémentés dans le worktree ; ils ne constituent pas encore une preuve d'exécution ou de persistance dans les hôtes VBA.

| Fonction | Commandes et limites |
| --- | --- |
| Transformations de code | `preview_code_edit` / `apply_code_edit` : remplacement littéral avec casse et mots entiers, remplacement lexical d'identifiants hors chaînes/commentaires, commentaires et indentation sur une plage explicite. SHA vérifié. Le remplacement lexical ne résout pas les portées ni les références COM ; ce n'est pas un renommage sémantique. |
| Annuler / refaire | `undo_code_edit` / `redo_code_edit` : journal des éditions VBAi `replace_lines` / `apply_code_edit`, par projet et module, limité à 50 entrées et 4 millions de caractères. Session courante uniquement, refus après une modification externe. La pile native VBE n'est pas utilisée. Une insertion ayant échoué déclenche une tentative de restauration avec relecture ; un rollback incomplet est signalé explicitement. |
| Recherche de définitions | `project_symbols` : modules, Sub, Function et Property du projet vivant, filtres et pagination, ligne et SHA pour naviguer. Les variables locales et la résolution sémantique ne sont pas couvertes. |
| Navigation / signets | `navigate_code` et `code_bookmark` : positions vérifiées par SHA, retour aux navigations VBAi et signets nommés en mémoire. Limites : 100 positions, 200 signets ; aucune persistance entre sessions. |
| Disposition des contrôles | `preview_form_layout` / `apply_form_layout` : alignement, tailles communes, distribution, centrage et grille, dans un même conteneur. Limites de géométrie, relecture et tentative de rollback. `set_form_tab_order` réordonne tous les contrôles directs du conteneur. La sauvegarde et le rendu natif restent à qualifier. |
| Listes persistantes imbriquées / multicolonnes | `set_form_list_initializer` accepte les chemins de Frames et Pages et soit `Items`, soit `Rows` (matrice rectangulaire, jusqu'à 64 lignes × 10 colonnes). Les colonnes doivent correspondre au `ColumnCount` existant. Génère `AddItem` et les affectations `List(row, column)` dans l'événement, sans mutation de la liste vivante du Designer. Contrôle de SHA, arbre, bloc géré et conservation du code utilisateur. Les listes liées à `RowSource` restent refusées. Exécution et persistance natives non encore vérifiées pour ces extensions. |
| Disposition des fenêtres | `window_layout` fournit une révision de géométrie et de liaison. `set_window_bounds` modifie une fenêtre autonome normale, avec relecture et restauration sur échec. `link_vbe_window` lie/détache un volet d'un cadre existant, vérifie les deux révisions et la nouvelle appartenance. Un échec natif partiel reste explicitement non vérifié ; un cadre détruit n'est pas présenté comme restauré. Persistance après redémarrage non vérifiée. |
| Connexion des compléments | `list_addins` fournit un `AddInVersion` ; `set_addin_connection` connecte/déconnecte un complément VBE déjà enregistré, vérifie identité et révision puis relit son état. Protection de VBAi par ProgId et GUID. Les échecs partiels ne déclenchent pas de nouvelle tentative ni de rollback implicite. Tests locaux passants ; validation d'un complément tiers réel encore à réaliser. |
| Listes liées Excel | `set_form_list_binding` remplace le bloc géré par une liaison `RowSource` vers une plage A1 de `ThisWorkbook`, sans écrire les cellules ni les propriétés du Designer. Vérification du chemin du classeur enregistré, des révisions, du nom de feuille, des limites de plage et du nombre de colonnes. Diff de code émis pour les deux commandes d'initialisation. Liaison au bon classeur, actualisation des valeurs et persistance vérifiées dans Excel 16.0 sur une liste imbriquée à deux colonnes. |
| Fractionnement du code | `set_code_split` demande explicitement un ou deux volets du module ciblé. Contrôle de SHA, de mode et du libellé natif de la commande 302, activation du module puis relecture de ses `CodePanes`. Idempotence et passage 1 → 2 → 1 vérifiés dans Excel 16.0. `list_commands` ignore les caractères de raccourci `&` pour la recherche mais conserve le libellé exact dans le résultat. |
| Contexte du chat | Scrutation incrémentale sur le thread STA, seulement lorsque le chat est visible et inactif : inventaire projets/modules/références et SHA des modules, à raison d'un module par tick de 1,5 s. Invalidation des références et actualisation des projets sans réaffecter implicitement une conversation. Le délai de détection croît avec le nombre de modules. Aucun abonnement natif `VBE.Events` n'est revendiqué. |

Validation locale : compilation de la solution sans erreur ni avertissement ; suite unitaire existante (375 tests) et tests additionnels de rollback / symboles passants ; chargement et modification de taille des 24 racines WinForms via `DesignSurface`. Ce dernier contrôle ne remplace pas une édition et sauvegarde manuelles dans Visual Studio.

**SOLIDWORKS : tests différés à la demande de l'utilisateur.** Aucun hôte n'a été lancé pour cette validation locale. Le thème VBE natif reste hors périmètre.

### Manques toujours ouverts — synthèse actualisée

Les sections chronologiques ci-dessous conservent les résultats antérieurs ; une limite ancienne peut être levée par une section plus récente. La couverture totale n’est pas acquise.

- Débogage : inventaire indépendant des points d’arrêt et du pointeur d’exécution, autres états/diagnostics et qualification des grands arbres/types de variables. Les sélections et retours de commandes ne remplacent pas cette preuve.
- Édition : renommage sémantique et qualification étendue du presse-papiers de contrôles/formats non textuels ; le presse-papiers Unicode du code est implémenté et prouvé. Restent aussi l’extension de la pile partagée de code aux projets avec UserForms et à plusieurs projets ouverts. L’historique ciblé des Designers est désormais implémenté et qualifié séparément. Les signets SQLite et l’annulation gérée par VBAi sont déjà implémentés et ont leurs preuves distinctes.
- Événements VBE : références, projets et composants désormais abonnés et prouvés pour les scénarios décrits plus bas. Le rechargement de composant reste non déclenché dans l’essai ; le texte des modules reste suivi par interrogation périodique.
- Options : mutation persistante et validation après réouverture encore en attente de l’autorisation spécifique décrite plus bas. Le thème natif VBE reste exclu.
- Formulaires : qualification systématique des propriétés/contrôles et événements, contrôles ActiveX tiers et interactions runtime étendues. Les cycles modal/modeless avec premier QueryClose annulé puis fermeture acceptée sont désormais qualifiés dans Excel. Les dispositions, listes imbriquées/multicolonnes et lancements modaux déjà prouvés ne couvrent pas toutes ces variantes.
- Fenêtres et extensions : autres DPI/géométries, persistance d’ancrage, boîte à outils, personnalisation des barres d’outils et qualification de leur disposition après redémarrage, connexion native d’un complément tiers. Le changement de vue simple/fractionné en mode conception et pause est désormais prouvé dans Excel.
- Projet : renommage et propriétés associés à un crash toujours gardés ; protection et autres fonctions natives de gestion restent à traiter avec une méthode vérifiée.
- Hôtes : sauvegarde native et lectures du débogueur propres à SOLIDWORKS restent des manques d’implémentation/qualification. Les tests SOLIDWORKS sont différés par l’utilisateur ; ils ne sont pas présentés comme réussis.

Dernière validation des listes : 381 tests unitaires passants, puis 9 tests de contrat LLM passants après ajout d'un contrôle explicite du schéma `Rows`. Les tests utilisent des doubles COM et ne prouvent pas le comportement natif.

Références de conception : [collections Microsoft Forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/objects-microsoft-forms) et [propriété List](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/list-property-microsoft-forms).

Fenêtres et compléments : compilation sans erreur ni avertissement ; tests ciblés passants. Le premier accès Excel a été refusé par COM. Après accord explicite de l'utilisateur, `AccessVBOM` a été activé uniquement pendant les essais puis restauré à son état initial (valeur absente). Les instances isolées ont été fermées/nettoyées.

### Disposition des fenêtres de documents

`editor_layout` et `arrange_editor_windows` exposent cascade, mosaïque verticale et horizontale. La mutation exige la version de disposition et le libellé exact de la commande native. Elle concerne toutes les fenêtres de documents visibles et non liées, y compris celles des autres projets ; les identités ambiguës sont refusées. L'identité repose sur le projet et le composant : VBE modifie les titres et recrée les objets Window lors du réagencement.

Essai Excel 16.0 avec trois modules : les trois commandes ont produit `Verified=true`, après relecture des identités et des géométries. Preuve locale : `artifacts/vbe-completion/native/window-arrangement-probe.json`, reproductible avec `tools/probes/Test-WindowArrangement.ps1`. L'instance isolée a été fermée et AccessVBOM restauré. Les designers, l'Explorateur d'objets, plusieurs projets simultanés et la persistance après redémarrage ne sont pas validés par cet essai.

### États et dimensions des fenêtres

`set_window_state` restaure, réduit ou agrandit une fenêtre exacte avec contrôle de `ExpectedWindowVersion`. Les volets liés doivent viser leur cadre. Une erreur native ou un état relu différent du résultat demandé reste non vérifié, même si la mutation a partiellement eu lieu. La relecture prend en compte le changement de titre des fenêtres de code.

Dans Excel isolé, un volet de code a effectué le cycle agrandir → restaurer → réduire → restaurer, avec les états natifs 2 → 0 → 1 → 0. `set_window_bounds` a ensuite relu exactement Left=40, Top=50, Width=600, Height=400. Preuve : `artifacts/vbe-completion/native/window-state-bounds-probe.json`, produite par le test de disposition. Les fenêtres d'autres types et la persistance après redémarrage restent à qualifier.

Référence : [propriétés VBIDE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model). `CodePaneView` est en lecture seule ; la mutation de cette propriété ne constitue pas une solution pour basculer entre procédure et module complet. `TopLine` est modifiable et exposé par `scroll_code_pane` avec ciblage précis des volets fractionnés.

### Défilement des volets fractionnés

`code_pane_layout` inspecte les volets déjà ouverts du module sans les activer. Chaque volet reçoit un jeton valable jusqu'au prochain appel de cette commande dans la session, une version de son affichage et le SHA du code. `scroll_code_pane` cible ce jeton, vérifie le projet/module, le mode conception ou pause attendu, la présence du volet, la version et le SHA. Le résultat n'est vérifié que si la ligne supérieure demandée est relue et si la sélection et le code sont conservés. Un clamp natif ou une erreur partielle reste non vérifié.

Excel 16.0 : une fenêtre fractionnée en deux volets a été positionnée respectivement aux lignes 30 et 70 ; les deux sélections sont restées `2:1:2:1` et le SHA du module est resté identique. Preuve : `artifacts/vbe-completion/native/code-pane-scroll-probe.json`, script `tools/probes/Test-CodePaneScroll.ps1`. Le mode pause et la vue procédure restent à qualifier, ainsi que les autres hôtes. La vue procédure/module complet reste une surface distincte non pilotée.

### Fermeture des fenêtres de documents

`close_vbe_window` a été essayé dans Excel sur un volet de code et un concepteur UserForm. Les deux ont disparu de `VBE.Windows` (`RemovedFromWindows`). Après réouverture par les objets VBIDE, le code complet était identique et le Label du formulaire conservait sa légende. Script : `tools/probes/Test-DocumentWindowClose.ps1` ; preuve : `artifacts/vbe-completion/native/document-window-close-probe.json`. Il s'agit d'un cycle dans un classeur jetable non enregistré, pas d'une preuve de persistance sur disque.

### Options natives : autorisation en attente

La lecture reste disponible. La revue automatique a refusé l'implémentation d'une mutation générique des options persistantes communes au VBE. Aucune modification de code ni de préférence n'a été appliquée par cette tentative. Une autorisation limitée aux réglages d'édition et de débogage, sans paramètres de sécurité, a été demandée. La mutation et sa validation après réouverture restent à réaliser.

### Dispositions MSForms : 42 cas conservés après réouverture

`align_centers` et `align_middles` complètent les alignements sur les bords : ils alignent respectivement les centres horizontaux et verticaux sur le premier contrôle sélectionné, en conservant leurs dimensions. Les résultats hors du conteneur sont refusés.

Le scénario `tools/tests/Test-ExcelFormLayouts.ps1` vérifie les 14 actions (`align_left/right/top/bottom/centers/middles`, `same_width/height/size`, `center_horizontal/vertical`, `snap_grid`, `distribute_horizontal/vertical`) sur trois boutons de tailles différentes, dans chacun des conteneurs UserForm, Frame et Page de MultiPage. Les invariants géométriques sont calculés indépendamment du plan retourné. Les 42 cas ont conservé leurs quatre dimensions après sauvegarde et réouverture d'un `.xlsm` jetable.

Preuve locale : `artifacts/vbe-completion/native/Vbai-EditorProbe-e1df02c750cf4dae97f07762139cffc5.xlsm.proof.json`, avec le classeur correspondant. Ce test qualifie ces conteneurs et CommandButton dans Excel ; il ne démontre pas chaque type ActiveX, chaque taille de sélection ni les autres hôtes. Les cas limites de centres fractionnaires et de dépassement sont couverts séparément par tests unitaires.

### Espacement des contrôles

Les actions `space_horizontal` / `space_vertical` imposent un écart exact (`Width`, en points ; zéro supprime l'écart). `increase_horizontal_spacing`, `increase_vertical_spacing`, `decrease_horizontal_spacing` et `decrease_vertical_spacing` augmentent ou diminuent chaque écart initial de `Width` points. Le contrôle le plus à gauche/en haut reste fixe ; les contrôles sont triés par position, puis chemin en cas d'égalité. Un écart négatif ou un résultat hors du conteneur est refusé avant écriture.

Le scénario de disposition inclut désormais ces six actions : 60 combinaisons (20 actions × 3 conteneurs), toutes relues après sauvegarde/réouverture dans Excel. Les six nouvelles actions ont été testées avec un écart/delta de 10 points. La suppression d'écart (zéro), les ordres de sélection différents et les refus sont couverts par les tests unitaires ; zéro n'a pas encore été essayé dans l'hôte.

Preuve : `artifacts/vbe-completion/native/Vbai-EditorProbe-6d8374a6cb6b4685b4f6ed6b1153ef9b.xlsm.proof.json`. Compilation sans erreur ni avertissement ; 21 tests ciblés passants. Cette preuve conserve les limites de types de contrôles et d'hôte du scénario précédent.

### Signets persistants par macro

`code_bookmark` conserve les signets des projets enregistrés dans une table dédiée de `%AppData%/CodexVBE/chat.db`. Le chemin absolu de la macro définit la portée, indépendamment du nom VBA du projet. Les noms de signets sont insensibles à la casse et le quota de 200 par macro est contrôlé dans une transaction SQLite. Les projets non enregistrés restent en mémoire avec une identité propre à l'objet projet ; leurs signets ne sont pas migrés automatiquement lors du premier enregistrement. Un déplacement ou Enregistrer sous crée une autre portée.

La navigation relit toujours le SHA avant de sélectionner le code. Essai Excel : signet créé via le nom du projet, retrouvé via son chemin après fermeture/réouverture du classeur et création d'un nouveau VbeSession, sélection native relue à `2:5`, puis refus après une édition externe. Script `tools/tests/Test-ExcelBookmarks.ps1`, base SQLite isolée du profil personnel. Preuve : `artifacts/vbe-completion/native/Vbai-EditorProbe-d0173ed8d4984ebf915cead8b379a852.xlsm.proof.json`.

Les tests SQLite réels couvrent aussi Unicode, casse, isolation entre chemins, suppression et nouvelle instance de navigation. L'historique Retour reste limité à la session. La même preuve ne valide pas un redémarrage complet de l'hôte ni SOLIDWORKS.

### Historique Aller / Retour / Suivant

`navigate_code` accepte `go`, `back` et `forward`. Les chemins et noms VBA d'un même projet enregistré sont normalisés avant de vérifier l'appartenance d'une position. Une nouvelle navigation efface la pile Suivant ; chaque pile est limitée à 100 positions, uniquement pour la session. La sélection n'est modifiée qu'après validation du SHA et une navigation refusée ne consomme pas l'entrée.

Essai Excel dans deux modules : aller par nom du projet, retour à la ligne/colonne d'origine, suivant par chemin du classeur, invalidation de Suivant après une nouvelle navigation. Un retour vers un module modifié a été refusé ; après restauration du code, la même entrée a permis de retrouver `2:5`. Script `tools/tests/Test-ExcelBookmarks.ps1`. Preuve : `artifacts/vbe-completion/native/Vbai-EditorProbe-5147b4e1395e49b9a2c306c45dc3ae81.xlsm.proof.json`. Cet historique VBAi reste distinct des commandes natives Définition/Dernière position du VBE.

### Navigation native Définition / Dernière position : observation

Le test `tools/probes/Test-NativeDefinition.ps1` a exécuté les commandes natives 939 (Définition) et 1822 (Dernière position) dans Excel, sur un appel `TargetMethod` entre deux modules temporaires. Définition a activé `DefinitionTarget`, ligne 2 colonne 1, sur `Debug.Print 42` ; la déclaration est ligne 1. Dernière position a réactivé `DefinitionCaller`, ligne 2 colonne 1. La commande ne garantit donc ni la sélection de la déclaration ni la conservation des colonnes de la sélection initiale.

Preuve locale : `artifacts/vbe-completion/native/native-definition-probe.json`. Le raccordement `native_code_navigation` est désormais présent : actions `definition`, `last_position` et `status` (Query=OperationId). Il planifie la commande sur le contexte UI, revalide mode/SHA/position avant exécution et conserve 20 résultats en session. Une seule opération peut être en attente. Le statut expose `CommandCompleted` et `NavigationObserved` ; il ne certifie pas la résolution sémantique. Une définition introuvable ou externe peut ouvrir un dialogue ou l'Explorateur d'objets. La gestion effective d'un dialogue modal dans le complément installé reste à tester. Le probe utilise maintenant le VbeSession courant avec une file STA de test : les deux navigations terminent, une seconde demande concurrente est refusée et une édition entre planification/exécution produit Failed sans exécuter la commande. Ce test ne prouve pas la boucle de messages du complément installé. Les cas externe/introuvable, les variables et propriétés restent à qualifier.

### Identification de la compilation et client du pont

`status` renvoie maintenant `AssemblyPath`, `AssemblyModuleVersionId`, `HostProcessId` et `ProcessBitness`, en plus du statut historique. Ces champs permettent de comparer le module effectivement chargé à la compilation attendue ; la version 0.1.0 seule ne distingue pas les builds.

L'inscription COM actuelle relue pointe sur `E:/Développement/AddIn/CodexVBA/bin/Debug/net48/CodexVBE.dll` (checkout main), pas sur la compilation de ce worktree. Elle n'a pas été modifiée. Le test via complément installé reste donc non exécuté pour les nouvelles commandes.

Le client `tools/Invoke-CodexVBE.ps1` borne désormais l'attente d'une réponse (30 secondes par défaut, paramètre `ResponseTimeoutSeconds`). Un délai expiré signale explicitement que la requête peut encore être en cours et ne déclenche pas de retry. Une fermeture sans réponse est distinguée du délai ; les erreurs de nettoyage ne masquent plus la cause initiale. `tools/tests/Test-BridgeClient.ps1` valide sur pipes locales les réponses normales, le silence (refus en environ 1 seconde pour le test) et la fermeture sans réponse. Ces tests de transport ne prouvent pas l'exécution dans Excel.

### Essai du complément enregistré : identité refusée

`tools/tests/Test-RegisteredExcelNavigation.ps1` sauvegarde trois valeurs CodeBase HKCU existantes (complément, version et contrôle de chat), les redirige temporairement vers la DLL du worktree puis les restaure et les relit dans `finally`. Il exige l'absence d'hôtes VBA ouverts et une option explicite. Le probe `-UseBridge` compare le PID et le ModuleVersionId avant toute navigation.

L'essai a été interrompu à ce contrôle : le pont répond seulement `{Version: 0.1.0, Connected: true}`, sans les nouveaux champs d'identité. Aucune navigation n'a donc été exécutée par le pont dans ce scénario. Les trois CodeBase temporaires ont été relus corrects pendant l'essai, puis restaurés. Un essai complémentaire a exporté la bibliothèque de types de cette compilation et redirigé temporairement sa valeur win64 également : le même ancien statut a été obtenu. Les quatre valeurs sont ensuite restaurées. Le processus de test et Excel ont le même propriétaire Windows ; la vue HKEY_CLASSES_ROOT relue par le probe montre bien le CodeBase temporaire. Ces observations écartent le simple oubli de la TLB ou un autre compte utilisateur, sans identifier encore la cause du chargement divergent. Diagnostic additionnel : `bridge-process-identity.json`. On ne conclut pas à une validation de la DLL du worktree. Diagnostic enregistré dans `artifacts/vbe-completion/native/bridge-identity-mismatch.json` ; sauvegarde des valeurs dans `temporary-registration-backup.json`.

### Chargement natif confirmé ; navigation installée encore en défaut

Le lancement direct `EXCEL.EXE /x` avec un classeur scratch identifié permet de charger la compilation attendue et de franchir les contrôles ModuleVersionId/PID du pont. Le démarrage par `New-Object -ComObject Excel.Application` donnait encore le format de statut ancien dans les essais précédents ; la cause interne de cette différence n'est pas établie. Le probe ferme le scratch sans enregistrer avant de créer le classeur jetable pour éviter deux projets VBAProject.

Le test installé révèle ensuite un écart réel : la commande Définition termine sans erreur, mais la sélection reste dans le module appelant (`NavigationObserved=false`). La relecture différée au passage UI suivant, le focus explicite sur le volet et l'emploi du contrôle retourné par FindControl n'ont pas résolu cet écart. Ces protections sont présentes ; elles ne constituent pas une preuve de navigation réussie. Le cas doit encore être diagnostiqué, notamment avec inspection des dialogues et des effets visibles. Les succès du probe COM externe ne remplacent pas ce test installé.

Le client vérifie désormais le PID réel du serveur via GetNamedPipeServerProcessId avant d'envoyer une requête ; Excel est bien le propriétaire du pipe testé. Les chemins COM et AccessVBOM sont restaurés après chaque essai.

### Navigation validée dans le complément installé

La capture indépendante `code_panes` a montré que VBE arrivait bien dans DefinitionTarget après la première relecture du service. Le suivi utilise désormais l'état `Observing` après le retour de la commande ; les lectures de statut peuvent actualiser la destination pendant deux secondes, sans attente sur le thread UI. Au-delà, une absence de déplacement reste non vérifiée. L'observation d'une sélection différente n'est toujours pas une certification sémantique.

Le scénario enregistré passe : lancement natif d'Excel, contrôle du PID serveur de pipe et du ModuleVersionId, Définition vers DefinitionTarget ligne 2 puis Dernière position vers DefinitionCaller ligne 2. Preuve : `artifacts/vbe-completion/native/installed-native-definition-probe.json`, qui contient l'identité de la DLL du worktree et les deux opérations terminées avec NavigationObserved=true. Cette preuve couvre bien le pont nommé et le contexte UI du complément chargé ; les cas introuvable/externe et les dialogues modaux restent non validés. Les quatre inscriptions COM temporaires et AccessVBOM sont restaurés après l'essai.

### Définition introuvable et dialogue modal : pont installé

Dans le complément chargé par Excel français, une définition introuvable affiche exactement « L'identificateur sous le curseur n'est pas reconnu ». Le filtre de `respond_debug_dialog` reconnaît désormais ce message exact, sans accepter un suffixe arbitraire ; il conserve la vérification du texte et de l'unicité du bouton avant action.

Le scénario installé a lu ce dialogue via `debug_dialog`, interrogé l'opération pendant son affichage, puis fermé le dialogue via `respond_debug_dialog` avec le bouton OK. La fermeture a été relue (`DialogClosed`). Après expiration des deux secondes d'observation, l'opération est terminée sans navigation observée ni résolution déclarée. Le pont est donc resté utilisable pendant ce dialogue précis.

L'essai a également montré que la commande COM peut retourner avant fermeture du dialogue : celui-ci peut rester ouvert pendant `Observing`, voire après `Completed`. Les descriptions d'outils ont été corrigées en conséquence. Preuve actualisée : `artifacts/vbe-completion/native/installed-native-definition-probe.json`, section MissingDefinition. Les variantes linguistiques de ce nouveau diagnostic, les définitions externes et les autres dialogues restent à qualifier.

### Définition externe : ouverture native de l'Explorateur

Le suivi inclut désormais la fenêtre active et la visibilité de l'Explorateur d'objets. Un changement de titre lié à l'état MDI n'est pas compté comme une navigation. L'apparition d'une fenêtre Type 2 l'est, même quand `ActiveWindow` renvoie le cadre principal Type 12. Si l'Explorateur était déjà visible et que le modèle objet ne révèle aucun autre changement, la navigation reste non observée plutôt que d'inventer la sélection d'un membre.

Dans le complément installé, Définition sur `Range` dans `Dim sample As Range` a ouvert l'Explorateur d'objets. Le service l'a observé et `vbe_windows` a relu la fenêtre Type 2 visible. Le scénario interne, le retour et le diagnostic d'identificateur introuvable passent aussi. Preuve actualisée : `installed-native-definition-probe.json`, section ExternalDefinition. `SemanticMemberVerified=false` : l'ouverture ne prouve pas encore la lecture de la bibliothèque/classe sélectionnée dans le volet.

### Vérifications natives Excel 16.0 — 28 septembre 2026

Script reproductible : `tools/tests/Test-ExcelEditorCompletion.ps1`, avec `-AssemblyPath`, `-OutputDirectory` et autorisation explicite `-AllowTemporaryVbaAccess` si nécessaire. Refuse de démarrer lorsqu'Excel est déjà ouvert. Invoque le `VbeSession` de la compilation courante contre les objets COM Excel réels ; ce n'est pas un test du transport du complément ni de l'UI du chat.

Preuve locale : `artifacts/vbe-completion/native/Vbai-EditorProbe-778514a5ed5c4e34ace3cdf126b98a4e.proof.json`, classeur jetable `Vbai-EditorProbe-778514a5ed5c4e34ace3cdf126b98a4e.xlsm`.

- Liste ListBox imbriquée dans un Frame, deux colonnes : génération, exécution, sauvegarde XLSM et réouverture ; valeurs `2|first|a"b|second|last` identiques. SHA du code du formulaire après réouverture : `147a735f40fc209cd356367c8be3fd0645d852256f93e7f824f620fa4bff9271`.
- Alignement gauche de deux boutons et ordre de tabulation de trois contrôles : relecture native puis relecture identique après réouverture.
- Édition avec commentaire, annulation, rétablissement puis restauration : SHA initiaux et édités vérifiés à chaque étape.
- Recherche de la définition `CheckGeneratedList` et navigation vers sa ligne : sélection relue par la commande native.
- Détachement puis rattachement de la fenêtre Exécution : résultats natifs vérifiés et ensemble des membres du cadre d'origine restauré.

- Conversion du même bloc d'items vers une liaison de feuille : résultat `2|bound|column|row2|value2` alors qu'un autre classeur portant la même feuille est actif. Une modification de la cellule source produit `2|changed|column|row2|value2`, conservé après nouvelle sauvegarde/réouverture. SHA final du code lié : `98b501c9d6f87c60d1a364240497bfa273b12c4a2c06c0d4d156c38f429f2078`. Les preuves du bloc d'items précédent correspondent à l'étape avant conversion.

- ComboBox à deux colonnes dans `MultiPage1/Pages/DataPage` : résultat natif `1|page row|page value` identique après sauvegarde/réouverture. SHA du code : `a7a9f28e3cb49a709ac3a790b619b3903fb3684fce47d06bc1db6645d0659e42`. Cette preuve couvre une page explicite et complète le cas Frame.

- Fractionnement via le service courant : deux volets obtenus, répétition sans basculement, retour à un volet relu. L'inventaire natif (`editor-command-inventory.json`) identifie aussi Cascade 1826, mosaïque verticale 2561 et horizontale 2562 ; leur invocation spécialisée et leur vérification restent à implémenter. Le changement vue procédure/module reste ouvert.

### Sondage de lecture des points d'arrêt

`tools/probes/Test-BreakpointCommandState.ps1` mesure `CommandBarButton.State` sur le bouton natif 51 dans un module jetable. Avant/après basculement et après déplacement du curseur vers une autre ligne, `State` reste à 0 ; le bouton 579 reste activé. Résultat : `artifacts/vbe-completion/native/breakpoint-command-state.json`. Ces états ne fournissent aucune preuve indépendante des marqueurs ; aucune commande d'inventaire n'est publiée sur cette base. Ce sondage ne prétend pas avoir vérifié un arrêt effectif en exécution.

Suite unitaire après cette extension : **394 tests passants**, aucun ignoré. [RowSource — référence Microsoft](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rowsource-property).

Ces preuves ne couvrent pas encore toutes les combinaisons : autres variantes de listes liées, tous les alignements, persistance de l'ancrage après redémarrage, connexion d'un complément tiers et tests SOLIDWORKS restent ouverts.

Références API : [LinkedWindows.Add](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-vba-add-in-object-model), [LinkedWindows.Remove](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/remove-method-vba-add-in-object-model) et [AddIn.Connect](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#connect).

Ces points ne sont pas marqués terminés par les tests locaux ; les détails historiques suivent.

## Carte de l'éditeur entier

| Surface VBE | Ce que CodexVBE couvre déjà | Manques précis à explorer |
| --- | --- | --- |
| [Fichier](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-menu) | Import, export et suppression de composants via VBIDE ; création de composants. `project_persistence_status` lit `VBProject.Saved` et l'état du classeur Excel correspondant ; `save_host_document` enregistre ce classeur avec vérification du chemin et de la version du projet. `save_host_document_as` a enregistré un classeur vierge en `.xlsm` neuf ; les SHA des éditions ont été retrouvés après réouverture en ciblant le projet par son chemin absolu. | Sauvegarde dans les autres hôtes, fermeture projet/VBE, impression et projets autonomes si l'hôte les propose. |
| [Édition](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/edit-menu) | Lecture et remplacement de lignes avec SHA, recherche littérale ou avec `*`/`?`, sélection/navigation du code ; suppression ciblée de contrôles. | Signets, remplacement avec portées et options, sélection multiple, indentation/commentaires, presse-papiers, annuler/refaire, complétion et informations de membres. Décider quelles actions d'éditeur doivent être primitives et lesquelles se déduisent des opérations de code. |
| [Affichage](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/view-menu) | Fenêtres et volets de code inventoriés ; formulaire ouvert ; Exécution, Variables locales, Espions, pile et Explorateur d'objets ouverts/lus partiellement. `show_vbe_window` réaffiche une fenêtre permanente masquée. Le menu **Affichage > Assistant CodexVBE** restaure déjà le panneau du chat. | Navigation Définition/position précédente, ordre de tabulation, boîte à outils, barre d'outils et personnalisation. Tester `show_vbe_window` sur les autres fenêtres permanentes des deux hôtes. |
| [Insertion](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/insert-menu) | Modules standard, classes, UserForms, contrôles natifs et stubs d'événement ; texte insérable par `replace_lines`. `create_procedure` ajoute une Sub, Function ou Property Get/Let/Set complète ; `replace_procedure` remplace seulement sa déclaration et son corps sous contrôle du SHA. Une Function et une Property Get ont été modifiées, compilées et retrouvées après réouverture dans Excel sans changer les commentaires ni les procédures voisines. `remove_procedure` a supprimé une Sub et un Property Get sous contrôle du SHA, en préservant le code voisin ; `insert_code_file` a ajouté une procédure depuis un chemin local explicite. Leurs SHA ont été conservés après réouverture d'un `.xlsm` jetable. | Designers et composants additionnels réellement disponibles ; autres encodages et positions d'insertion à qualifier. |
| [Format](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-menu) | Position/taille individuelles des contrôles, ZOrder avant/arrière prouvé sur deux Labels. | Opérations sur sélection multiple : alignement, espacement, taille commune, centrage, grille, groupes, ajustement au contenu et déplacement d'un rang dans l'ordre Z ; lecture du résultat et persistance. |
| [Débogage](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu) | Points d'arrêt natifs, compilation, espions, pas à pas, jusqu'au curseur, instruction suivante. | Inventaire indépendant des points d'arrêt et du pointeur d'exécution ; détails dans la section Débogueur. |
| [Exécution](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/run-menu) | Démarrer, poursuivre, interrompre, réinitialiser et exécuter une ligne dans Exécution essayés. `run_sub` exécute par nom une Sub sans paramètre d'un module standard ; la cellule modifiée a été relue dans Excel. | Mode conception par projet, lancement d'un UserForm via commande dédiée, dialogue Macro, procédures avec paramètres et variantes de projets autonomes propres à l'hôte. |
| [Outils](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tools-menu) | Références sélectionnées lues/ajoutées/retirées ; `Description`, `HelpFile` et `HelpContextID` écrits et relus dans Excel ; `Instancing` de classe conservé après réouverture ; `list_signing_certificates`, `read_project_signature_dialog` et `sign_project` exposés. `read_vbe_options` lit les quatre onglets natifs Éditeur, Format, Général et Ancrage sans enregistrer. Sur un `.xlsm` contenant un vrai module, la signature avec le certificat déjà associé a été enregistrée, les parties `vbaProjectSignature*.bin` constatées et [`Workbook.VBASigned`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbasigned) relu à `true` après réouverture. | Mutation vérifiée des options, protection, gestion des macros, contrôles supplémentaires et paramètres du projet. La première sélection d'un certificat passe par la fenêtre protégée « Sécurité Windows » : l'utilisateur doit la confirmer ; le complément gère le [dialogue VBE](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/digital-signature-dialog-box) et vérifie le nom retenu, mais ne pilote pas les boutons du sélecteur Windows ni ne valide la chaîne de confiance. Hors Excel, la sauvegarde et la vérification durable restent propres à l'hôte. |
| [Compléments](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-ins-menu) | CodexVBE se charge dans Excel et SOLIDWORKS après réparation de l'installation ; `list_addins` lit les compléments enregistrés dans le VBE et leur état `Connect`. Dans Excel, CodexVBE a été relu connecté. | Chargement/déchargement via le gestionnaire, diagnostic d'installation par hôte et réouverture fiable du panneau ; tester plusieurs compléments et SOLIDWORKS. |
| [Fenêtre](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/window-menu-commands) | `vbe_windows` et `code_panes` relisent fenêtres, visibilité, géométrie et sélection ; `focus_vbe_window` active une fenêtre visible exacte. `close_vbe_window` a masqué l'Explorateur d'objets, puis `show_vbe_window` l'a réaffiché dans Excel. `window_linkage` a relu le cadre et quatre fenêtres liées de l'Explorateur de projets. Chat hébergé à droite avec disposition restaurée. | Fractionner le code, cascade/mosaïque, changement des liaisons et de l'ancrage ; tester la fermeture d'un volet de code et d'un designer, premier placement droit sur un profil vierge. |
| [Aide](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/help-menu) | `vbe_environment` lit la version du moteur VBE et les comptes des collections principales. | Accès contextuel à l'aide et identification plus détaillée de l'hôte, sans faire passer l'ouverture d'une aide pour la lecture de son contenu. |
| [Menus contextuels](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shortcut-menu) et barres | `list_commands` inventorie les CommandBars, chemins, identifiants et état activé ; quelques commandes sont exécutées de façon spécialisée. | Cartographier par contexte Projet/Code/Formulaire/Fenêtre, car l'état de ces commandes dépend de la sélection. Qualifier leurs effets avant d'exposer une exécution générique. Aucun raccourci clavier ni clic à coordonnées. |

## Modèle objet transversal

| Objet ou relation | État | Manque |
| --- | --- | --- |
| `VBE.VBProjects` et projet actif | Liste et mode, propriétés partielles ; `project_properties` lit l'état `Protection` en lecture seule (`vbext_pp_none` dans Excel jetable) ; cible par nom unique ou `FileName` absolu en cas de collision ; état de sauvegarde du projet et du classeur Excel, sauvegarde Excel vérifiée après réouverture. | Identité des projets sans chemin et de même nom, réglage natif de la protection, sauvegarde et cycle de vie dans les autres hôtes. |
| `VBComponents`, `CodeModule`, `CodePane` | Composants et code lus/édités ; procédures, sélection sur une ou plusieurs lignes et ajout, remplacement ou suppression ciblés d'une procédure complète. La sélection multiligne a été relue dans le volet Excel après un second appel. | Méthodes VBIDE non encore exposées selon leur intérêt, vue complète du code fractionné et options de l'éditeur. |
| `References`, bibliothèques et Explorateur d'objets | Références sélectionnées et métadonnées COM paginées ; Explorateur ouvert. | Résolution transitive, références cassées, objets/membres de projet dans l'Explorateur, recherche et navigation structurées natives. |
| `Windows`, `LinkedWindows`, `CommandBars`, `AddIns` | Fenêtres/commandes lues, focus exact d'une fenêtre visible vérifié, fermeture/réouverture d'une fenêtre permanente confirmée ; `LinkedWindowFrame.LinkedWindows` relu ; compléments VBE et état `Connect` inventoriés. | Modifier l'ancrage, vérifier sa persistance, fermer les autres types de fenêtres, charger/décharger les compléments et suivre les événements des barres de commandes. |
| `VBE.Events` | Aucun abonnement aux événements du modèle objet VBE n'a été trouvé dans `src` ; la connexion actuelle passe par `IDTExtensibility2`. | Brancher, si l'hôte les expose, les changements de projet/composants/références et de commandes utiles pour invalider les instantanés et maintenir le contexte du chat. |

## Débogueur

| Surface | État actuel | Travail restant |
| --- | --- | --- |
| Points d'arrêt | Basculement natif, arrêt effectif et effacement de deux points dans une macro jetable prouvés. | Lire la liste et les lignes de tous les points d'arrêt de manière indépendante. Sans cette lecture, `toggle_breakpoint` et l'effacement global gardent une vérification limitée. |
| Instruction suivante | `show_next_statement` et déplacement dans une même procédure essayés ; `set_next_statement` refuse une procédure différente d'après la sélection du volet de code. | Identifier indépendamment le pointeur d'exécution. Une sélection déplacée manuellement n'est pas une preuve de sa position ; vérifier les autres cas d'erreur natifs. |
| Exécution et pas | Exécuter, poursuivre, interrompre, réinitialiser, pas détaillé, principal, sortant et jusqu'au curseur essayés dans Excel. Compilation et exécution native d'une procédure temporaire prouvées dans SOLIDWORKS par une sortie sur disque ; point d'arrêt atteint en mode pause ligne 5 et Pas à pas principal confirmé ligne 6. | Vérifier les autres pas et la reprise dans SOLIDWORKS ; la fenêtre Exécution de cet hôte n'a pas livré son texte par la méthode actuelle. |
| Variables et expressions | Variables locales, Espions, pile, Exécution, espion express et modification de variable via Exécution essayés dans Excel. | Dans SOLIDWORKS 2019, Variables locales est visible pendant l'arrêt mais UI Automation ne fournit aucune ligne ; la fenêtre Exécution ne fournit pas de document texte. Adapter la lecture à cet hôte, puis qualifier les grands arbres d'objets, types particuliers et autres langues. Une expression VBA peut avoir des effets de bord. |
| Compilation et erreurs | Diagnostic de compilation et dialogue d'erreur d'exécution natifs lus dans Excel ; option d'arrêt sur erreur lue. | Détails de diagnostic au-delà du message/texte sélectionné, état d'exception une fois le dialogue fermé, et cas d'erreur particuliers. Une compilation sans dialogue n'a pas de preuve binaire indépendante. |

## Concepteur de formulaires

| Surface | État actuel | Travail restant |
| --- | --- | --- |
| Types et propriétés | Les 14 contrôles MSForms natifs ont été créés ; 771 descripteurs et 110 choix d'énumération recensés. Plusieurs propriétés scalaires, `Font` et `Picture` sont relues après écriture. | Qualifier chaque propriété réellement modifiable par type, valeur, persistance après sauvegarde/réouverture et effet dans le formulaire exécuté. Les 621 candidats d'écriture issus des descripteurs ne sont pas 621 écritures prouvées. |
| Propriétés à risque | Certaines écritures sont bloquées après échec natif ou crash corrélé dans Excel. | Comprendre la cause avec une méthode sûre avant tout nouvel essai ; ne pas lever les gardes sur une simple indication `ReadOnly=False`. |
| Listes ComboBox/ListBox | Lecture de l'état vivant présente. `AddItem` direct en conception a perdu ses lignes après sauvegarde/réouverture. `set_form_list_initializer` est exposée au LLM et a conservé les items au runtime après réouverture d'un `.xlsm` pour des listes non liées à une colonne. | Couvrir les listes liées, multicolonnes et contrôles imbriqués avec essais distincts ; vérifier l'initialisation dans les autres hôtes. |
| Arbre et disposition | Frame, Page, Tab, contrôles imbriqués, dimensions, suppression et ordre visuel de deux Labels essayés. | Qualifier les autres combinaisons de conteneurs/contrôles, l'ordre Z au-delà de ce cas, la boîte à outils et les contrôles ActiveX tiers réellement hébergés. |
| Événements | Catalogue des interfaces source COM et création de stubs par `CreateEventProc` présents. | Couvrir les événements utilisables par VBA que le seul catalogue COM n'annonce pas, et vérifier les signatures/contexte de chaque contrôle dans l'hôte. |

## Autres surfaces de l'éditeur

| Surface | État actuel | Travail restant |
| --- | --- | --- |
| Projets et code | Composants, références, lecture/édition SHA, création, remplacement, suppression et inventaire de procédures, recherche littérale et à jokers, insertion d'un fichier local avec provenance ; portée `Instancing` d'une classe pilotable ; `run_sub` cible une procédure par son nom et le SHA du module. L'état `Saved` du projet et du classeur est lu ; en Excel, `remove_procedure` et `insert_code_file` ont conservé leurs SHA après sauvegarde et réouverture. | Renommage de projet désactivé après crash corrélé ; remplacement avancé, cas d'erreur et compatibilité documentaire par hôte. |
| Signature VBA | `sign_project` exige un projet enregistré et un certificat de signature de code explicite par empreinte ; le complément écarte la collision native de l'ID 746 avec « Supprimer le composant ». En Excel, il enregistre le classeur exact après la signature et relit `VBASigned`. | Confirmer manuellement la première sélection dans « Sécurité Windows », qualifier la sauvegarde et la vérification dans SOLIDWORKS, et vérifier la confiance/validité cryptographique du certificat au-delà de l'état signé d'Excel. |
| Explorateur d'objets | Ouverture native et métadonnées des références sélectionnées présentes. | Catalogue structuré des bibliothèques/classes/membres dans le complément, navigation et recherche vérifiées ; l'essai UI Automation exploratoire est externe au complément. |
| Fenêtres et menus | Inventaire des fenêtres, volets de code et CommandBars présent ; focus, fermeture et réouverture exacts exposés ; cadre et membres de `LinkedWindows` lisibles ; hébergement du chat à droite et restauration de cette disposition validés dans Excel. | Pilotage borné des commandes non encore spécialisées, mutation d'ancrage et premier placement droit sur profil vierge. |
| Hôtes | Les fonctions ci-dessus ont été principalement essayées dans Excel 64 bits visible. Dans SOLIDWORKS 2019, le complément est connecté ; lecture des projets/fenêtres/compléments/options, sélection multiligne, fenêtres liées et Explorateur d'objets vérifiés. Le projet `test.swp` jetable a permis les cycles d'export/import/suppression de module, classe et UserForm avec Label et propriétés, puis compilation, exécution, arrêt sur point, pas principal et reprise ; les composants et le SHA du code initial ont été restaurés. | Qualifier les autres écritures et commandes de débogage, la lecture des volets natifs et la sauvegarde SOLIDWORKS. Les cycles temporaires ne prouvent pas une persistance sur disque dans cet hôte. |

Les preuves et limites par essai sont détaillées dans [vbe-coverage.md](vbe-coverage.md), [forms-property-coverage.md](forms-property-coverage.md) et [llm-integration.md](llm-integration.md). Cette liste est un suivi de couverture, pas une déclaration que toutes les fonctions VBE ont été inventoriées par une API exhaustive.


## 28 septembre — lecture native de l’Explorateur d’objets

- `read_object_browser` est disponible par le bridge et par l’agent, hors du thread UI VBE. Lecture seule des sélections, valeurs et description native ; aucune ouverture ou sélection implicite.
- Le parcours cible les HWND des contrôles et évite d’énumérer tous les éléments des bibliothèques COM. Les titres français/anglais sont pris en charge ; une fenêtre non identifiée est signalée indisponible. Les erreurs de lecture sont conservées.
- Validation réelle Excel x64 via `Test-RegisteredExcelNavigation.ps1` : navigation interne et dernière position, diagnostic inconnu, ouverture de l’Explorateur et lecture des sélections/description. Preuve : `artifacts/vbe-completion/native/object-browser-read.json`, identité du binaire dans `installed-native-definition-probe.json`.
- Limite observée : Définition sur `Range` ouvre l’Explorateur mais la sélection native reste `Classes <globales>` et la description `<Toutes bibliothèques>`. Aucune résolution de `Range` n’est revendiquée. Recherche/sélection de classe ou membre encore à couvrir.
- Inscriptions COM temporaires restaurées et vérifiées ; instance Excel fermée, AccessVBOM remis à son état initial absent. Aucun test SOLIDWORKS.


## 28 septembre — sélection native d’une classe et d’un membre

- `select_object_browser` expose `ObjectName` et `Procedure` facultatif au bridge et à l’agent. La fenêtre doit déjà être ouverte. Sélection UIA par libellé exact, refus des doublons, notification native puis relecture des sélections. Les membres sont recherchés avec le préfixe exact de la classe pour éviter de sélectionner un membre de la classe précédente.
- Excel x64 réel : `Range` puis `Address`, sélection native `Classes Range` / `Membres de 'Range' Address`, description `Property Address(...) As String` et `Membre de Excel.Range` vérifiées. Preuve : `artifacts/vbe-completion/native/object-browser-selection.json`.
- Une classe inexistante est refusée et la description précédente est conservée (`object-browser-guards.json`). Erreurs et résultats partiels ne déclenchent pas de répétition automatique.
- 407 tests unitaires réussis. Validation native par le binaire identifié dans `installed-native-definition-probe.json`, inscriptions COM restaurées après le test.
- Restent notamment le choix explicite de bibliothèque, la recherche et la pagination des listes natives. En cas de noms ambigus, la commande refuse actuellement la sélection. Le parcours anglais est implémenté mais seul le VBE français a été validé en réel.


## 28 septembre — recherche et pagination de l’Explorateur

- `list_object_browser` lit les listes `classes`, `members`, `libraries`, filtre les libellés par sous-chaîne sans tenir compte de la casse et fournit Offset/Limit (50 par défaut, 200 maximum), TotalMatches et HasMore. Aucun clic ni modification des listes. Les listes non identifiables/vides restent explicitement indisponibles.
- Excel x64 réel : recherche de la bibliothèque `Excel` (un résultat), deux pages distinctes de deux membres de `Range`, 199 membres observés. Preuve : `artifacts/vbe-completion/native/object-browser-lists.json`. La pagination des classes est implémentée mais pas encore validée dans l’hôte.
- `select_object_browser` accepte désormais une bibliothèque `Context`. **Changement de bibliothèque non validé** : le ComboBox natif est désactivé dans l’instance observée, y compris après activation de la fenêtre. Le garde refuse cette opération ; la preuve porte `LibraryChange.Verified=false`. Aucun contrôle désactivé n’est forcé. La sélection de classe/membre sans changement de bibliothèque reste validée.
- Compilation réussie, 10 tests de contrats réussis (arguments et bornes contrôlés avant l’accès natif). La restauration COM et la fermeture Excel ont été confirmées après le scénario.


## 28 septembre — bibliothèque native : blocage modal identifié et résolu

Les premières conclusions sur un ComboBox désactivé sont précisées : la fenêtre principale VBE était désactivée par un diagnostic distinct, `Impossible d'aller à 'Range' qui est caché`. L’ouverture de l’Explorateur ne signifiait pas la fermeture de ce diagnostic. La lecture Win32 a identifié cet ancêtre, puis `debug_dialog` a fourni le message exact (`external-definition-dialog.json`).

- Le diagnostic français observé est reconnu par un motif strict et borné. La réponse exige toujours le message exact et un bouton natif identifié. Les suffixes et sauts de ligne inattendus sont refusés par les tests.
- `select_object_browser` refuse maintenant toute navigation si sa fenêtre ou un ancêtre est désactivé. Les lectures restent disponibles. Le scénario réel prouve ce refus avant fermeture explicite du diagnostic.
- Le choix de bibliothèque utilise une correspondance exacte, vérifie les fenêtres natives actives, sélectionne l’index, notifie le parent et relit le libellé UIA. Aucune fenêtre n’est réactivée de force. Comportement des messages : [documentation Microsoft ComboBox](https://learn.microsoft.com/en-us/windows/win32/controls/about-combo-boxes).
- **Validé sous Excel x64 : `Context=Excel`, `ObjectName=Range`, `Procedure=Address`**, bibliothèque et sélections relues, description `Membre de Excel.Range` conservée. `object-browser-lists.json` porte désormais `LibraryChange.SelectionObserved=true`. La limitation de bibliothèque notée dans la section précédente est donc levée dans ce scénario.
- La navigation Définition sur `Range` reste un échec sémantique natif (élément caché), même si elle ouvre le volet. Le chemin explicite de sélection fonctionne après fermeture du diagnostic. La variante anglaise de ce nouveau diagnostic reste à qualifier.


## 28 septembre — bibliothèque seule et catalogue des classes validés

- `select_object_browser` accepte `Context` seul ; `Procedure` exige toujours `ObjectName`. La relecture finale contrôle aussi la bibliothèque quand elle est demandée.
- La sélection utilise maintenant l’action UI Automation du ComboBox actif, après vérification des ancêtres Win32. Le précédent essai fondé sur les messages ComboBox est remplacé par ce chemin complet, validé dans le même hôte.
- Cycle réel : Excel → VBA seule, description native `Library VBA`, recherche insensible à la casse `collection` donnant `Collection` et `MatchCollection`, deux pages distinctes de classes, refus d’une bibliothèque inexistante sans perdre VBA, retour à `Excel.Range.Address`. Preuve : `artifacts/vbe-completion/native/object-browser-library-cycle.json` ; binaire/hôte identifiés par `installed-native-definition-probe.json`.
- Compilation sans erreur ; 10 tests de contrats réussis après modification des arguments. L’essai natif confirme la restauration COM et la fermeture d’Excel. SOLIDWORKS et le thème VBE natif restent exclus des essais.


## 28 septembre — exécution native dédiée des UserForms

- `run_form` cible le Designer du UserForm exact et planifie la commande native 186 sur le thread VBE. Requiert le projet, le formulaire, le mode conception, le SHA du code, la TreeVersion et le libellé exact de la commande. Vérifie à nouveau les révisions et l’identité COM avant l’exécution ; refuse une autre commande UserForm déjà en attente.
- `form_run_status` lit l’opération par projet et identifiant. Les 20 dernières opérations restent en session. `Queued`, `Running`, `CommandReturned` et `Failed` décrivent la commande, pas une preuve d’affichage. `RuntimeVerified=false` reste explicite. Le modèle doit lire les diagnostics ou obtenir une preuve d’exécution ; aucun retry implicite. Les politiques lecture seule et mode Plan refusent l’exécution.
- Excel x64 réel : formulaire jetable avec `UserForm_Activate` écrivant `VBAi form executed` en A1 du classeur de test puis `Unload Me`. Valeur relue et commande terminée ; aucun dialogue restant. Deux demandes préalables portant un SHA ou une TreeVersion périmés ont été refusées, sans écrire le marqueur.
- Preuve : `artifacts/vbe-completion/native/native-run-form.json`, binaire identifié dans `installed-native-definition-probe.json`. 407 tests unitaires réussis. Les formulaires restant modaux, erreurs d’événements et autres hôtes restent à qualifier ; cette preuve ne couvre pas toute interaction runtime des contrôles.


## 28 septembre — UserForm modal et erreur d’événement

- Formulaire réellement visible (`ThunderDFrame`) avec titre unique, marqueur écrit par `Activate`, statut d’opération consulté pendant l’affichage, fermeture native de ce seul formulaire de test et `QueryCloseMode=0` relu dans le classeur jetable. `CommandReturned` était déjà présent pendant l’affichage : il ne signifie donc pas que le formulaire est fermé. Preuve : `native-run-form-modal.json`.
- Le test utilise `VBComponent.Properties.Item("Caption").Value`. L’affectation directe sur `Designer.Caption` n’avait pas produit le titre runtime attendu dans ce scénario ; cette tentative n’est pas une preuve de mutation de Caption.
- Une erreur non gérée `Err.Raise 5` dans `UserForm_Activate` ouvre une fenêtre dont le titre est **Microsoft Visual Basic**, différent des titres précédemment recherchés. Le lecteur de diagnostics, les réponses et les gardes de compilation reconnaissent désormais ce titre exact dans le processus hôte. La validation du message exact et du bouton est conservée.
- Preuve réelle : diagnostic `Erreur d'exécution '5'`, marqueur de l’événement relu, bouton exact `&Fin`, fermeture confirmée puis Mode=2. Fichier : `native-run-form-error.json`. Aucune réussite runtime n’est déduite du seul retour de la commande ; `RuntimeVerified` reste false dans son statut.
- Scénario enregistré dans `tools/probes/Test-NativeDefinition.ps1`, identité du binaire dans `installed-native-definition-probe.json`. Inscriptions COM restaurées et instance Excel fermée. Les autres événements, annulation par QueryClose, formulaires modeless et interactions runtime des contrôles restent à qualifier.


## 28 septembre — vue procédure / module complet

- `set_code_view` accepte `procedure` ou `module`, avec jeton de volet, WindowVersion, SHA, mode attendu et ligne source valide. Refuse un volet remplacé, une révision périmée ou une cible différente après le focus. Une demande portant déjà sur la vue courante ne bascule pas la vue.
- `CodePaneView` est [en lecture seule dans VBIDE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#codepaneview). UIA et MSAA n’ont exposé aucune action pour les deux boutons dans l’hôte testé. Sondage et capture : `tools/probes/Test-CodePaneView.ps1`, `code-view-window.png`, `code-view-msaa.txt`.
- Les boutons appartiennent à un contrôle enfant `ObtbarWndClass`. Le complément identifie le code window exact dans son processus, vérifie les ancêtres actifs et l’emplacement du contrôle par rapport à la barre de défilement. Les messages souris ciblent ce HWND avec des coordonnées locales calculées à partir de ses dimensions ; aucun déplacement global du pointeur. Une géométrie inconnue est refusée.
- Excel x64, complément installé : cycle 1 → 0 → 1, sélection `2:1:2:1` et SHA conservés. Les versions périmées sont refusées et les demandes répétées sont des no-op vérifiés. Preuve : `installed-code-view-cycle.json`, identité du binaire dans `installed-native-definition-probe.json`.
- Compilation réussie et 407 tests unitaires réussis. Limites explicites : le changement effectif exige actuellement une fenêtre non fractionnée ; mode pause, autres DPI/géométries, persistance et autres hôtes restent à qualifier. Cette commande ne prétend pas couvrir ces cas.


## 28 septembre — vues des fenêtres fractionnées et volets COM détruits

- Le sondage réel montre que la vue procédure/module est commune à la fenêtre : les deux `CodePaneView` changent ensemble. `set_code_view` annonce désormais cette portée, restitue `BeforePanes`/`AfterPanes` et vérifie le mode de vue et le SHA de chaque volet vivant.
- Les versions de tous les volets inspectés sont conservées. Une sélection changée dans le second volet invalide une demande ciblant le premier, avant le clic natif. Preuve : `installed-split-code-view-guard.json`.
- Cycle installé Excel : fractionnement relu dans une requête distincte, changement procédure en ciblant le premier volet, changement module en ciblant le second, vérification des deux vues et de leurs SHA, puis fusion et relecture d’un seul volet valide. Preuve : `installed-split-code-view-cycle.json` ; script `Test-RegisteredExcelNavigation.ps1`.
- Défaut natif rencontré : après fusion, `CodePanes` peut encore énumérer le volet détruit, dont chaque propriété renvoie `DISP_E_BADCALLEE (0x80020010)`. `code_pane_layout` signale cette entrée dans `UnavailablePanes` et ne lui attribue aucun jeton ; le volet vivant reste utilisable. Le comptage du fractionnement utilise le même traitement ciblé. Les autres erreurs COM ne sont pas masquées. Relevé brut : `installed-unsplit-panes.json`.
- 410 tests unitaires réussis, dont conservation de l’identité du module vivant, diagnostic du volet détruit et propagation des autres erreurs COM. La restriction précédente aux fenêtres non fractionnées est levée pour ce scénario Excel. Mode pause, autres DPI et persistance restent à qualifier.


## 28 septembre — changement de vue en mode pause

- Procédure jetable : écrit 10 en A5, exécute `Stop`, puis incrémente A5 après reprise.
- Complément installé Excel : vue procédure puis module sur un volet, fractionnement en mode pause, vue procédure puis module sur les deux volets. Les quatre cas conservent Mode=1 et A5=10 ; chaque vue et SHA sont relus.
- Commande native Continuer ensuite : Mode=2, A5=11 et SHA identique. Cette preuve montre la conservation de l’arrêt et la reprise effective dans ce scénario ; elle ne constitue pas un inventaire indépendant du pointeur d’exécution.
- Preuve : `artifacts/vbe-completion/native/installed-break-code-view-cycle.json`, produite par `tools/tests/Test-RegisteredExcelNavigation.ps1`. Le binaire testé est celui identifié dans `installed-native-definition-probe.json`. Aucun changement des options de débogage n’est effectué.
- Mode pause qualifié pour cette fenêtre simple/fractionnée Excel. Les autres DPI/géométries et la persistance restent ouverts.

### Événements natifs des références — 28 septembre 2026

Le chat s’abonne désormais à `VBE.Events.ReferencesEvents` du projet sélectionné : ajout et suppression invalident le contexte, dont le rafraîchissement est différé jusqu’au prochain passage idle visible. Le changement de macro retire les deux anciens abonnements ; la fermeture du chat arrête le timer et retire les abonnements. Un échec de connexion conserve la surveillance périodique, toujours nécessaire pour les modifications du texte des modules.

Les IID/DISPID ont été vérifiés dans l’interop Microsoft VBE 15 installé. Preuve Excel isolé : `tools/probes/Test-ReferenceEvents.ps1`, résultat `artifacts/vbe-completion/native/reference-events-native.json` : ajout, suppression, ancien projet silencieux après changement de portée, nouveau projet notifié et absence de notifications après Dispose. Le test utilise un compteur .NET, car les callbacks de script PowerShell ne mettaient pas à jour son compteur. Il valide la classe de production contre Excel via COM ; le rafraîchissement visuel du chat n’est pas une preuve de cet essai.

Compilation sans erreur ni avertissement et 412 tests unitaires passants. Les événements de composants/projets et de modification du code ne sont pas ajoutés par ce changement. Référence : [événements documentés du modèle VBE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/events-visual-basic-add-in-model).

### Historique natif du code — pile partagée et diff réels

`native_code_history_state` capture les SHA de tous les modules, les commandes Annuler/Rétablir et une `HistoryVersion`. `native_code_history` exige cette révision (`ExpectedProjectVersion`), le libellé exact, le mode conception et exécute une seule commande native 128/129. Tous les modules sont relus ; les changements réels alimentent `CodeEdited` et les diff du chat. La pile native est distincte du journal `undo_code_edit`/`redo_code_edit` de VBAi. Aucune sauvegarde implicite.

**Constat natif déterminant : la pile est partagée entre modules.** Activer un module ne garantit pas qu’Annuler le modifiera : Excel a annulé l’insertion dans le dernier module édité et activé celui-ci. La première approche ciblant seulement le module actif a été remplacée avant qualification. Le contrat final décrit donc explicitement cette portée partagée. Il refuse plusieurs projets ouverts pour respecter la macro liée au chat, et les projets contenant un UserForm tant que la pile Designer n’est pas qualifiée. Ces limites restent des manques, pas une couverture complète.

Preuve reproductible : `tools/probes/Test-NativeCodeHistory.ps1`, résultat `artifacts/vbe-completion/native/native-code-history.json`. Deux modules : annuler vide le dernier édité, rétablir restaure son SHA ; l’autre module affiché reste inchangé. Le test vérifie également le refus d’une ancienne révision, d’un mauvais libellé, d’un autre projet ouvert et d’un UserForm. L’appel passe par `LlmVbeTools`, avec deux notifications `CodeEdited` sur le module effectivement modifié ; il ne constitue pas une capture du rendu visuel du chat. Le test utilise une instance Excel isolée et restaure AccessVBOM.

Compilation sans erreur/avertissement ; 412 tests unitaires passants, dont les refus en mode Plan et lecture seule. Si aucune modification n’est observée ou si une erreur de relecture survient, l’outil ne prétend pas connaître le contenu de la pile ; il demande une nouvelle lecture et interdit de déduire qu’une répétition automatique serait sûre. L’historique natif ne donne pas de garantie de persistance. Référence de comportement : [menu Édition du VBE](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/edit-menu).

### Historique natif ciblé des UserForms

`native_form_history` appelle `Designer.UndoAction()` ou `Designer.RedoAction()` sur la UserForm exacte, en mode conception, après contrôle de `ExpectedTreeVersion` (arbre et états CanUndo/CanRedo relus). Cette voie cible le Designer directement ; elle n’utilise pas la pile de code partagée ni une activation de fenêtre pour choisir la cible. Les changements de propriétés faits par COM ne créent pas systématiquement d’entrée dans l’historique natif.

Le résultat expose `DesignerChanges` (chemin canonique, propriété, avant/après), l’arbre relu et les erreurs natives. Le comparateur traite aussi les contrôles imbriqués, ajouts/suppressions, empreintes d’images et membres de police lorsqu’ils sont disponibles. CanUndo/CanRedo/CanPaste ne sont pas des changements de conception. Une propriété en erreur est exclue des différences et les compteurs `ReadErrorsBefore/After` signalent la lecture partielle : `Verified` confirme des changements observés, pas l’exhaustivité des propriétés.

Qualification Excel : `tools/probes/Test-NativeFormHistory.ps1`, preuve `artifacts/vbe-completion/native/native-form-history.json`. Un alignement natif de deux boutons est annulé (SecondButton.Left 55 → 100) puis rétabli (100 → 55), alors qu’un autre Designer dans un autre projet est actif ; l’arbre de cet autre formulaire reste inchangé. Une révision périmée et une action indisponible sont refusées. Une suppression native de deux contrôles est ensuite annulée (0 → 2 contrôles). Dans cette instance, CanRedo est alors faux : le rétablissement est refusé sans nouvelle suppression, plutôt que déclaré réussi.

La compilation réussit sans erreur/avertissement ; 415 tests unitaires passent, y compris différences imbriquées, lectures en erreur et politiques Plan/lecture seule. Cette preuve ne couvre ni tous les ActiveX, ni la persistance après sauvegarde, ni un rendu visuel spécifique des différences de Designer dans le chat. Le thème VBE natif reste exclu ; SOLIDWORKS reste non testé. Référence : [UndoAction Microsoft Forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/undoaction-method).

### Presse-papiers Unicode du code

`read_code_clipboard`, `copy_code`, `cut_code` et `paste_code` exposent le presse-papiers Windows et des plages explicites de code (lignes/colonnes à partir de 1, fin exclusive). Copier/couper exigent une plage non vide ; coller accepte une position d’insertion et le module vide (1:1 → 1:1). Les outils ne déplacent pas le curseur. Le presse-papiers est global à Windows, et son outil de lecture indique de l’utiliser uniquement pour une demande qui le concerne.

La coupe et le collage exigent le mode conception et le SHA courant. Le collage exige aussi la `Version` du presse-papiers précédemment lu (`ExpectedClipboardVersion`) ; la lecture vérifie son numéro de séquence Windows et son empreinte. La copie est relue avant de supprimer la source. Une erreur d’édition après la copie laisse explicitement le texte dans le presse-papiers. Les changements passent par `replace_lines`, son rollback d’erreur et le journal VBAi ; les diff sont transmis au chat, y compris lors d’un échec si la relecture révèle une mutation partielle. Un collage identique n’écrit pas le code. Si VBE normalise le texte, la relecture signale que le résultat exact demandé n’est pas vérifié.

Preuve Excel/Windows : `tools/probes/Test-CodeClipboard.ps1` et `artifacts/vbe-completion/native/native-code-clipboard.json`. Copie de « alpha », coupe, collage avec remplacement, annuler/refaire, refus d’un presse-papiers périmé et collage de commentaires accentués dans un module vide ; cinq notifications de diff observées via `LlmVbeTools`. Le test sauvegarde les formats chaîne/MemoryStream du presse-papiers uniquement en mémoire et les restaure avec relecture exacte ; les formats qu’il ne sait pas sauvegarder sont refusés avant mutation. Excel est fermé et AccessVBOM restauré.

423 tests unitaires passants couvrent notamment plages multilignes, erreurs de copie, révisions périmées, mode pause, presse-papiers sans texte, caractères de contrôle et collage identique. Limites : texte Unicode jusqu’à 1 Mi caractères, fins de ligne normalisées CRLF ; aucune prise en charge du transfert de contrôles Designer, d’images ou de formats riches. La qualification native utilise des caractères accentués compatibles avec cet Excel ; elle ne garantit pas la conservation de tous les caractères Unicode par le stockage VBA de chaque hôte. Les colonnes désignent des positions de caractères, pas des coordonnées visuelles.

### Notifications natives des projets et composants

Le chat conserve désormais les abonnements des collections `VBProjects` (ajout, suppression, renommage, activation) et `VBComponents` du projet sélectionné (ajout, suppression, renommage, rechargement), en plus des références déjà intégrées. Les IID/signatures ont été lus dans l’interop Microsoft installé et les points de connexion confirmés dans Excel. `FindConnectionPoint` peut retourner null sans erreur pour une interface absente : la connexion vérifie donc le pointeur et son IID avant abonnement.

Les callbacks ne lisent pas l’objet COM reçu, qui peut déjà être invalide après suppression ; ils invalident seulement le contexte géré. Les abonnements de composants changent avec la macro, et tous sont retirés à la fermeture du chat. L’échec partiel retire les abonnements réussis et conserve la surveillance périodique. La fermeture d’un projet force le rafraîchissement des portées avant de relire son ancien sélecteur, pour éviter qu’une lecture échouée bloque indéfiniment la mise à jour. La conversation n’est jamais réaffectée implicitement à une autre macro.

Preuves : `tools/probes/Test-VbeConnectionPoints.ps1` / `vbe-event-connection-points.json` et `tools/probes/Test-CollectionEvents.ps1` / `collection-events-native.json`, dans `artifacts/vbe-completion/native`. Les événements de composants ajout/renommage/suppression donnent les compteurs 1/2/3 ; après changement de portée, l’ancien projet reste à 3 et le nouveau porte le compteur à 4. Les projets ouvrent/renomment/ferment avec notifications ; après Dispose, les compteurs composants/projets restent 4/7. La collection COM doit être conservée comme objet .NET dans le harnais, sans dépliage PowerShell en éléments.

Compilation sans erreur ni avertissement ; 425 tests unitaires passants. Les notifications sont vérifiées contre Excel isolé, pas par capture du rafraîchissement visuel du chat. L’événement ItemReloaded est raccordé mais n’a pas été déclenché par cette qualification. Il n’y a pas d’événement de modification du texte ajouté : la vérification incrémentale des SHA reste nécessaire. Excel est fermé et AccessVBOM restauré après l’essai.

### UserForms en exécution : modal, modeless et QueryClose

`Test-RegisteredExcelNavigation.ps1 -Scenario FormLifecycle` charge temporairement la DLL du worktree dans Excel et exécute `Test-FormLifecycle.ps1` par le bridge nommé du complément. Le test contrôle PID, chemin et MVID de la DLL chargée. ShowModal est modifié par `set_form_property` avec révision, puis relu avant lancement ; un objet COM de propriétés peut devenir invalide après déchargement et n’est pas réutilisé dans le test.

Preuve `artifacts/vbe-completion/native/installed-form-lifecycle.json` :

- **Modeless** : ShowModal=false, fenêtre Excel activée, UserForm visible et événement Activate observé ; état VBE Mode=2 pendant l’affichage.
- **Modal** : ShowModal=true, fenêtre Excel désactivée, UserForm visible et Activate observé ; état VBE Mode=0 pendant l’affichage.
- Dans les deux cas, `form_run_status` reste accessible et indique CommandReturned sans prétendre vérifier la durée de vie de la fenêtre. La première fermeture native provoque QueryClose (CloseMode=0), annulé par le code ; le formulaire reste visible. La seconde fermeture est acceptée, compteur=2, retour au mode conception et SHA du module inchangé.

**Mode=2 ne prouve donc pas qu’aucune UserForm n’est encore affichée.** La description de `debug_state` précise cette limite. Le lancement n’impose plus un mot français/anglais dans le libellé : il contrôle l’ID natif 186, son état activé et le libellé exact précédemment lu. Des tests unitaires couvrent des libellés allemands, espagnols et japonais ; ces langues ne sont pas présentées comme qualifiées dans des hôtes natifs.

429 tests unitaires passants ; compilation sans erreur. Après le test, Excel est fermé, AccessVBOM est absent comme avant l’essai, les trois CodeBase et la TypeLib sont restaurés. Les interactions détaillées avec les contrôles runtime, le déclenchement d’autres événements et les variantes ActiveX restent à couvrir. SOLIDWORKS reste différé et le thème VBE natif exclu.

### Inventaire des fenêtres UserForm en exécution

`read_runtime_forms` est un outil de lecture asynchrone, disponible par le bridge et dans le chat. Il énumère les fenêtres visibles de premier niveau `ThunderDFrame` du PID hôte : handle, titre, activation, propriétaire natif et dimensions. Il ne lit ni les contrôles ni leurs valeurs et ne demande pas le modèle COM VBA, ce qui permet sa lecture pendant une boucle modale. Le résultat indique explicitement que l’identité du projet n’est pas résolue ; un titre de fenêtre n’est pas traité comme une preuve de macro ou de nom de UserForm.

Qualification dans `installed-form-lifecycle.json` (DLL/MVID du bridge contrôlés) : zéro fenêtre runtime avant lancement malgré un Designer ouvert ; exactement la fenêtre attendue après lancement modal et modeless ; même handle présent après QueryClose annulé ; zéro après fermeture acceptée. Ces assertions sont intégrées à `Test-FormLifecycle.ps1`. Aucun dialogue Designer n’a été compté comme formulaire runtime dans cet essai.

Limites conservées : fenêtres cachées et classes autres que ThunderDFrame exclues, 64 entrées maximum avec indication de troncature, handles transitoires, aucune association automatique à une macro. La lecture ne prouve pas l’exécution des événements d’initialisation ; la preuve de ces événements reste le marqueur du test. `form_run_status` recommande désormais aussi cette lecture. 429 tests unitaires passants ; compilation sans erreur ; instance Excel et inscriptions COM temporaires nettoyées/restaurées.

### Visibilité des barres d’outils VBE

`list_toolbars` inspecte les barres normales (`Type=0`) de `VBE.CommandBars` : nom exact, visibilité, activation, origine intégrée/personnalisée, protection, position et géométrie. Les erreurs de géométrie sont distinctes des erreurs d’état. `WindowVersion` couvre le nom, le type, la visibilité, l’activation et la protection ; il ne promet pas une révision de la disposition complète.

`set_toolbar_visibility` applique show/hide au nom exact, après contrôle de cette révision, puis relit Visible. Les noms ambigus, menus et menus contextuels sont refusés. Montrer une barre désactivée est refusé sans activer la barre ni modifier sa protection. Une opération déjà satisfaite ne réécrit pas Visible. L’opération concerne l’UI globale du VBE ; aucun code VBA n’est modifié.

Preuve Excel isolé : `tools/probes/Test-ToolbarVisibility.ps1` et `artifacts/vbe-completion/native/native-toolbar-visibility.json`. La barre Standard est masquée puis réaffichée, avec révision initiale et géométrie restaurées (Position=1, RowIndex=5, Left=1, Top=25, Width=664, Height=26 dans cet essai). Une barre temporaire effectue également show/hide, puis est supprimée. Une ancienne révision est refusée et une demande sans changement n’écrit rien. Ce résultat vérifie les propriétés COM natives ; la persistance après redémarrage n’est pas revendiquée.

Compilation sans erreur ; 431 tests unitaires passants. Excel est fermé et AccessVBOM restauré. L’ancrage, la personnalisation des boutons, la boîte à outils MSForms et la persistance restent des sujets distincts. Références : [CommandBar.Visible](https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.visible) et [MsoBarType](https://learn.microsoft.com/en-us/office/vba/api/office.msobartype).

### Ancrage des barres d’outils VBE

`set_toolbar_position` accepte gauche, haut, droite, bas et flottant sur une barre normale identifiée par son nom exact. `ExpectedToolbarLayoutVersion`, lu dans `list_toolbars`, couvre état et géométrie ; une lecture incomplète ou une révision périmée empêche la mutation. Les protections natives contre le déplacement et les ancrages sont respectées. Le masquage respecte également NoChangeVisible. Aucune protection n’est désactivée par ces outils.

Preuve Excel isolé : `tools/probes/Test-ToolbarDocking.ps1` et `artifacts/vbe-completion/native/native-toolbar-docking.json`. Les cinq positions ont été appliquées et relues sur une barre temporaire. Une ancienne révision après déplacement est refusée, une barre masquée reste masquée après ancrage, et NoChangeDock empêche le déplacement. La barre temporaire est supprimée, Excel fermé et AccessVBOM restauré à son absence initiale. Cette preuve utilise la session de production via COM ; elle ne constitue pas une qualification du bridge installé.

438 tests unitaires passants, compilation réussie. Les coordonnées flottantes précises, les rangées, les effets sur les barres voisines et la persistance après redémarrage restent non qualifiés. L’ancrage natif peut réorganiser les barres voisines ; le résultat expose cette limite et ne revendique aucune persistance.

### Placement précis des barres d’outils

`set_toolbar_placement` expose deux actions sans changer le mode d’ancrage : `float` avec `ToolbarLeft` / `ToolbarTop` explicites en pixels entiers, et `row` avec `RowIndex` positif. Nom exact, révision complète et protections sont contrôlés avant écriture. La position, la visibilité et les valeurs relues doivent correspondre pour obtenir Verified=true. Une normalisation native ou une erreur partielle reste explicitement non vérifiée, avec état après opération ; aucune nouvelle tentative automatique.

Le probe `Test-ToolbarDocking.ps1` qualifie désormais (250,180) sur la barre temporaire flottante puis RowIndex=2 sur la barre ancrée masquée. Les deux résultats sont vérifiés dans `native-toolbar-docking.json`. 439 tests unitaires passent. Excel est fermé et AccessVBOM restauré. La persistance, les variantes DPI/écrans et la disposition des barres voisines restent ouvertes.

Sémantique native : [Left](https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.left), [Top](https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.top), [RowIndex](https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.rowindex).

### Presse-papiers natif du Designer

`form_clipboard_state` expose l’arbre et la sélection native du Designer exact, une révision combinant ces deux états, la révision Windows du presse-papiers et CanPaste. `native_form_clipboard` applique Copy/Cut/Paste au Designer explicitement choisi, sans commande globale dépendant du focus. Les états périmés sont refusés. Ces outils sont disponibles dans le chat ; les mutations suivent sa politique d’édition. Le contenu binaire du presse-papiers n’est pas transmis au modèle.

Preuve `tools/probes/Test-FormClipboard.ps1`, `artifacts/vbe-completion/native/native-form-clipboard.json` : bouton sélectionné détecté ; copie ; collage dans un autre Designer pendant que la source a le focus ; légende relue ; coupe de la source pendant que la cible a le focus ; états périmés refusés. Le presse-papiers initial est sauvegardé en mémoire puis restauré et comparé format par format. Excel est fermé et AccessVBOM restauré. La session de production est appelée par COM depuis le probe, sans qualification du bridge installé pour ce scénario.

**La coupe COM ne peuple pas la pile Undo dans cet essai.** CanUndo=false est lu, la demande d’annulation est refusée sans mutation, puis le recollage explicite depuis le presse-papiers inchangé rétablit un contrôle. Il ne s’agit pas d’une preuve de restauration de toutes ses propriétés/positions. Les résultats exposent les différences observées, les erreurs de lecture et ClipboardContentVerified=false. Ils ne promettent pas une fidélité binaire ni un rollback automatique.

Compilation réussie et 440 tests unitaires passants. Restent la sélection programmatique, les sélections multiples/imbriquées, les différents contrôles et formats, la fidélité complète après sauvegarde/réouverture et la récupération indépendante du presse-papiers. La sémantique de Copy dépend de l’objet cible : [documentation MSForms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copy-method-microsoft-forms).

### Sélection et presse-papiers par conteneur

`select_form_controls` remplace la sélection d’un conteneur par une liste de noms directs exacts (Items, 0 à 64 ; vide pour désélectionner). Il valide la révision de sélection/arbre et tous les noms avant écriture, puis relit Selected. La propriété native InSelection est utilisée sans clic, raccourci ou changement de focus. Une erreur partielle est accompagnée de l’état relu.

`form_clipboard_state` et `native_form_clipboard` acceptent maintenant ParentPath : absent pour la racine, chemin canonique vers une Frame ou une Page sinon. La révision inclut ce chemin. Les sélections sont indépendantes : Selected de la racine n’énumère pas les enfants sélectionnés d’une Frame/Page. La copie et la coupe sont donc invoquées sur le conteneur explicite.

Le probe `Test-FormClipboard.ps1` n’utilise plus Select All pour préparer ses opérations. Preuves natives dans `native-form-clipboard.json` : sélection racine programmée ; une puis deux sélections dans une Frame ; révision périmée refusée ; nom inexistant refusé sans perte de sélection ; copie Label/CommandButton de Frame vers Page, légendes relues ; coupe et récupération par recollage ; désélection du conteneur. Sauvegarde xlsm/réouverture : deux enfants présents dans chaque conteneur, légendes cibles conservées. Le gestionnaire Click source est inchangé et le module cible reste sans code, avant et après sauvegarde. Cela ne prouve pas la fidélité de toutes les propriétés, images ou contrôles tiers.

Compilation sans avertissement ni erreur ; 440 tests unitaires passants. Presse-papiers restauré et comparé, Excel fermé, AccessVBOM restauré à son absence initiale. La récupération indépendante du presse-papiers, les autres types de contrôles et la qualification exhaustive des propriétés restent ouvertes. Le thème natif VBE reste exclu et les essais SOLIDWORKS différés.

### Récupération de coupe indépendante du presse-papiers courant

La coupe appelle désormais Copy puis capture immédiatement les formats lisibles avant Cut. Elle refuse la suppression si la capture échoue, si MS Forms Bag manque, si le presse-papiers change durant la capture ou si les données dépassent 8 Mio. Une erreur peut laisser la copie dans le presse-papiers ; le résultat expose l’erreur et l’état relu. Les huit dernières captures sont conservées uniquement dans la session VbeForms, avec identité de l’objet formulaire et chemin de conteneur. Les captures évincées/fermées ne sont pas promises comme récupérables.

`restore_form_clipboard` utilise DesignerClipboardRecoveryId retourné par la coupe, la même identité de formulaire/conteneur et les révisions courantes. Il republie les formats capturés et compare leurs octets/textes relus. Il ne colle pas et ne supprime pas d’éventuels doublons : l’agent doit inspecter l’arbre puis demander un collage explicite. Cela ne constitue pas une restauration automatique de géométrie ni une annulation native.

Preuve dans `native-form-clipboard.json` : remplacement du presse-papiers par du texte après coupe racine et imbriquée ; restauration depuis la capture ; recollage ; refus d’un autre formulaire ; sauvegarde/réouverture avec contrôles/légendes attendus et code événementiel inchangé. Le presse-papiers initial est finalement restauré, Excel fermé, AccessVBOM remis à son absence initiale. `MS Forms CLSID` est annoncé mais illisible via WinForms et figure dans OmittedFormats ; les formats lisibles suffisent pour ces contrôles testés, sans preuve de fidélité universelle.

442 tests unitaires passent, dont indépendance des octets copiés et refus des captures absentes/surdimensionnées. Compilation réussie. La restauration de toutes les propriétés, images et types ActiveX, l’interface visuelle de rollback et la persistance des captures entre sessions restent ouvertes.

### Récupération de coupe avec géométrie et ordre de tabulation

`recover_form_cut` consomme une capture de coupe de cette session. Il vérifie la même identité de formulaire/conteneur, les révisions actuelles, l’absence de différence lisible depuis la coupe et l’absence des contrôles à recréer. Il republie le presse-papiers capturé, colle une seule fois puis rétablit les noms attendus, les rectangles des contrôles coupés et l’ordre TabIndex de tous les contrôles directs, y compris ceux restés en place. L’identifiant ne permet pas de recommencer après une tentative de collage, même partiellement échouée. `restore_form_clipboard` reste disponible séparément pour inspection/récupération manuelle.

Preuve native actualisée dans `native-form-clipboard.json` : bouton racine récupéré à (30,30) ; deux contrôles imbriqués aux géométries distinctes récupérés autour d’un Label resté en place ; ordre de tabulation exact ; remplacement du presse-papiers sans perte de récupération ; refus de rejouer la récupération ; refus d’un changement de légende intermédiaire. Les rectangles et TabIndex sont comparés directement puis de nouveau après sauvegarde/réouverture du xlsm. Le code événementiel reste inchangé.

Aucune différence résiduelle sur les propriétés comparables dans ces deux cas. Cela ne prouve pas une fidélité complète : 5 erreurs de lecture de propriétés après récupération racine et 16 pour l’arbre imbriqué restent exposées, ainsi que le format illisible MS Forms CLSID. Le résultat distingue RestoredNamesGeometryAndTabOrder de FullPropertyFidelityVerified=false. Les changements de propriétés illisibles ne peuvent pas être détectés par le contrôle des modifications intermédiaires.

442 tests unitaires passants, compilation réussie, git diff --check sans défaut. Presse-papiers initial restauré, Excel fermé et AccessVBOM restauré. Il reste notamment à brancher une action de rollback visible dans le chat et à qualifier les autres propriétés/types de contrôles ; l’objectif de couverture complète reste ouvert.

### Carte de récupération dans le chat

Une coupe effectuée par les outils LLM émet désormais une carte de contrôles coupés avec une action « Restaurer les contrôles coupés ». Le clic utilisateur relit les révisions, valide le périmètre du chat puis appelle recover_form_cut sans nouvelle confirmation. La disponibilité vérifie la capture en mémoire et l’identité du formulaire ; le bouton est désactivé pendant une réponse, après tentative/restauration ou expiration. Les tokens et l’objet propriétaire ne sont jamais sérialisés dans SQLite : l’historique rouvert et les branches de conversation n’acquièrent pas une nouvelle autorité de récupération.

Preuve native `Test-FormClipboard.ps1` : événement émis par LlmVbeTools lors de la coupe, rendu du composant réel, événement Click du bouton, restauration dans Excel et désactivation. Il s’agit d’un clic programmatique du composant réel, pas d’un test manuel dans le chat du complément installé. Géométrie, tabulation et sauvegarde/réouverture restent contrôlées par ce probe.

`Test-FormRecoveryCard.ps1` vérifie trois états désactivés et leurs tooltips en clair/sombre. Captures `artifacts/vbe-completion/ui/form-recovery-Light.png` et `form-recovery-Dark.png` inspectées. Les huit libellés sont traduits localement dans tous les catalogues existants. 443 tests unitaires passent, dont la non-persistance de l’autorité de récupération et la complétude des catalogues. La couverture VBE totale reste ouverte selon la synthèse en tête du document.
