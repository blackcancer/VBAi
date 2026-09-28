# Qualification native des fonctions VBE

État du 28 septembre 2026. Les preuves ci-dessous portent sur des scénarios précis ; elles ne transforment pas les fonctions absentes ou bornées de [functional-extensions.md](functional-extensions.md) en une couverture universelle. Le thème natif VBE reste exclu.

## Lot de complétion IDE après PR #6

Les tests vivent dans le projet VSTest, sous `Integration/Hosts/Excel`, avec leurs helpers sous `Infrastructure/Hosts`. Chaque fixture ouvre Excel et son VBE de façon visible, travaille sur son seul classeur jetable et le ferme.

| Contrat | Preuve Excel obtenue | Limite restante |
| --- | --- | --- |
| Renommage public intermodules | Deux modules relus, appel qualifié modifié, exécution VBA donnant 42 dans une cellule, annulation exacte séparée des deux modules | Pas de liaison complète des membres de classe/interfaces/callbacks/clients externes ; pas d'atomicité globale |
| Ajustement UserForm | Dimensions VBIDE/InsideWidth/InsideHeight relues indépendamment ; enfants inchangés ; étendue de défilement relue | Arrondi supérieur au pixel selon le DPI ; conteneurs imbriqués pas tous qualifiés |
| Explorateur de projets | HWND SysTreeView32 réel, nœuds UIA lus avec leur identité | Variantes hôtes et manipulations expand/collapse à qualifier |
| Boîte à outils | Page « Contrôles » relue via le fournisseur MSAA canonique ; identité/version stables ; refus de sélection sans mutation | Lecture des pages seulement ; mutations et boutons de palette non qualifiés |
| Protection | Lecture avec Cancel ; verrouillage depuis fichier secret, confirmation des contrôles, sauvegarde XLSM puis fermeture/réouverture ; VBProject.Protection=1 relu indépendamment | Aucun déverrouillage de projet protégé ; Word/SOLIDWORKS NOT_RUN ; simple fermeture du dialogue ne prouve pas la persistance |
| Options Format/Ancrage | Modifications de cases, réouverture et vérification ; restauration de la version complète des préférences | Choix de police/palettes et autres langues pas tous qualifiés |
| Valeurs de procédure | SAFEARRAY vector/matrix, Null, paramètres optionnels, bornes retournées 1/-2 et 0/0 ; compteur Excel confirmait une invocation par appel et aucune invocation lors du polling ; code inchangé | Excel uniquement ; pas d'objets COM/classes/ByRef/tableaux typés/ParamArray |

Preuves : `artifacts/vbe-completion/native-matrix/matrix.trx`, `fit-final/fit.trx`, `protection-qualified/protection.trx`. Les passages exploratoires échoués restent distincts du passage global final. Le premier rapport contient un échec de protection corrigé ensuite : il ne doit pas être présenté comme entièrement vert.

L'exploration des couleurs du TextPattern sur un code comportant un point d'arrêt connu et en pause renvoie BackgroundColorAttribute=NotSupported. Cette voie ne prouve ni l'inventaire des points d'arrêt ni le pointeur d'exécution. La typelib MSForms inspectée expose Selected/InSelection et CanUndo/CanRedo, mais aucun identifiant/appartenance de groupe ; Group/Ungroup ne sont pas qualifiés par un arbre inchangé.

Le cycle standalone `.swp`, les adaptateurs Word/PowerPoint et l'aide CHM ont des scénarios locaux mais restent **NOT_RUN** dans ces hôtes. Le SIP VBA Microsoft reste soumis à l'autorisation déjà demandée ; aucun enregistrement cryptographique n'a été effectué.

### Boîte à outils : observation native sans mutation

La fenêtre réelle utilise une classe `F3 MinFrame`, avec des enfants `F3 Server`. La sonde externe UI Automation ne livre aucun onglet ni bouton de palette. Dans l'add-in, UIA projette plusieurs représentations du même onglet, sans état de sélection, ainsi que le bouton de fermeture de la fenêtre. La lecture privilégie donc le fournisseur MSAA canonique, qui expose une liste d'onglets et la page « Contrôles », sélectionnée. `accSelect` et l'action par défaut « Passer à » retournent normalement ; la page unique étant déjà sélectionnée, cela ne qualifie pas un changement de page.

`MSForms Toolbox.ShowPopup()` sans arguments affiche le menu exact, mais les commandes 2591 à 2596 restent désactivées : le menu visible ne fournit pas le contexte natif de l'onglet. Aucun onglet n'a été créé, renommé, supprimé ou déplacé. Les boutons de palette demeurent inaccessibles. Les sondes ont fermé leurs seules instances Excel possédées.

Preuves exploratoires : `artifacts/vbe-completion/native/toolbox-survey.json`, `toolbox-msaa.json`, `toolbox-activation.json` et `toolbox-typelib.txt`. La matrice `artifacts/vbe-completion/toolbox-canonical/matrix.trx` contient 21 tests réussis, dont la lecture réelle, la stabilité de la version et le refus d'action sans changement. `read_navigation_surface` expose ces pages au LLM ; aucune commande de personnalisation n'est prétendue disponible.

## Qualifications natives précédentes

## Résultats exacts

| Point | Preuve obtenue | Ce qui reste |
| --- | --- | --- |
| 1. Appels paramétrés depuis le complément installé | Excel isolé, PID et MVID du complément vérifiés. Sub avec chaîne française contenant des guillemets, Double, Boolean et Null ; Function appelée avec 21, résultat 42 relu indépendamment dans la feuille. | Objets et retour structuré restent hors contrat ; variantes d’hôtes non qualifiées. Les arguments nommés ont une qualification complémentaire ci-dessous. |
| 2. Préférences et barres persistantes | Six cases Éditeur, tabulation aux bornes 1 et 32, trois choix d’interruption sur erreurs, deux cases de compilation : **13 essais** avec réouverture, relecture et restauration du hash de toutes les options. Barre et bouton natif ID 186 restaurés après fermeture/reprise d’Excel ; identité exacte du contrôle conservée par Copy. Position flottante (320, 220) également restaurée et relue après redémarrage. | Préférences natives anglaises et autres hôtes non exécutés. La restauration par profil concerne uniquement les barres VBAi explicitement persistantes ; elle ne personnalise pas les menus/boîtes à outils. |
| 3. SWP, rechargement et signatures | Aucun nouvel essai SOLIDWORKS effectué dans ce lot. | **En attente du fichier jetable `.swp` autorisé**. Save/SaveAs, rechargement, SHA et signature après persistance restent non qualifiés. Le contrôle hors ligne d’un certificat ne valide pas le digest de la signature VBA. |
| 4. UserForms, propriétés, événements et ActiveX | 14 types MSForms : **34 mutations/restaurations** (Left, Enabled et Caption quand exposée), relecture native indépendante ; huit Caption absentes/non lisibles pour leur type. **60 dispositions** sauvegardées et réouvertes sur UserForm, Frame et page MultiPage. Cycle modal/modeless avec QueryClose annulé puis accepté, état du parent et source vérifiés. | Les neuf MSComctl présents au catalogue sont tous **HOSTING_REFUSED**, zéro contrôle restant, avec « sujet non approuvé ». Aucune politique de confiance modifiée. Les autres propriétés et événements ne sont pas tous qualifiés ; catalogue de propriétés ≠ setter natif validé. |
| 5. Grands arbres, langues, DPI et ancrage | Tableau Excel de **1 000 Long**, expansion/repli, sept lignes scalaires (Double, Currency, Date, String, Boolean, Variant/Null, Long), reprise et marqueur 3000 ; source inchangée. **52 captures** de quatre fenêtres dans 13 cultures ; **8 captures** des quatre fenêtres sur deux écrans réels, à 96 DPI. | Arbres COM, espions volumineux, débogueur SOLIDWORKS et autres langues natives VBE non qualifiés. Les captures WinForms ne valident pas la qualité linguistique complète, le DPI 125/150/200 % ou l’ancrage natif VBE. Certaines chaînes restent en anglais par repli du catalogue. |

Ce tableau décrit les scénarios exécutés. Aucun de ces résultats ne prouve une couverture universelle de toutes les combinaisons VBA/COM.

## Corrections révélées par ces scénarios

- `Temporary=false` ne conservait pas effectivement la barre au redémarrage d’Excel : profils SQLite privés dans `%LocalAppData%/VBAi/VbeToolbars/<HOST>.sqlite`, fusion transactionnelle, bornes et suppression par identité.
- Ajouter un bouton par ID Office créait « Macros » au lieu du contrôle VBE « Exécuter » : copie du contrôle natif existant, ID/type/possession vérifiés, libellé de restauration dépendant du contexte.
- Les libellés français réels de complétion, info express, compilation sur demande et largeur de tabulation différaient de la liste autorisée. Le champ de tabulation partage aussi son nom avec une étiquette ; seul le type éditable est sélectionné.
- Une séquence de mutations MSForms suivie de l’ajout d’un contrôle a provoqué une corruption de heap Excel (`0xc0000374`). Le membre fautif n’a pas été isolé individuellement. Les propriétés scalaires usuelles utilisent désormais un dispatch COM typé ; la matrice de 34 mutations passe avec cette correction. Les autres gardes restent actives.
- UIA exposait deux fois chaque enfant du tableau. La capture retire uniquement les observations identiques par texte complet et chemin ; **1 000 doublons retirés, 1 000 valeurs conservées**, dernière valeur 3000. Des valeurs/contextes/chemins distincts sont conservés. Le contrat reste `UIAExposedRowsOnly`.

### Gardes MSForms conservées

Écriture refusée pour `ComboBox.ColumnCount`, `TextBox.ScrollBars`, `SpinButton.Min/Max/Value/Delay/SmallChange`, `ToggleButton.Value`, et `Label.Cancel` sans setter utilisable. `_Font_Reserved` est un membre COM réservé indisponible. Ces restrictions ne sont pas levées par les essais des propriétés usuelles.

## Reproduire et consulter les preuves

Les scénarios refusent les hôtes VBA déjà ouverts, identifient leur propre processus et utilisent un classeur jetable. L’enregistrement COM temporaire et AccessVBOM sont restaurés ; aucune macro utilisateur n’est exécutée. Le wrapper exige les opt-in explicites `AllowTemporaryRegistration` et `AllowTemporaryVbaAccess`.

- `tools/tests/Test-RegisteredExcelNavigation.ps1`, scénario `FunctionalExtensions` : appelle `tools/probes/Test-RegisteredFunctionalExtensions.ps1` ; preuves finales dans `artifacts/native-qualification/registered-final/` (rapport, matrice des options, propriétés, ActiveX et nettoyage).
- Même wrapper, scénario `DebuggerTree` : appelle `tools/probes/Test-RegisteredDebuggerTree.ps1` ; `artifacts/native-qualification/debugger-tree-final/registered-debugger-tree.json`.
- Scénario `FormLifecycle` : `artifacts/native-qualification/form-lifecycle/installed-form-lifecycle.json`.
- `tools/tests/Test-ExcelFormLayouts.ps1` : preuve dans `artifacts/native-qualification/form-layouts/`.
- `tools/tests/Test-MultilingualWindows.ps1` : captures dans `artifacts/localization/multilingual-windows/`.
- `tools/tests/Test-WinFormsDisplayProfiles.ps1` : mesures/captures dans `artifacts/native-qualification/displays/` ; DPI du processus et écran effectivement utilisés enregistrés.

Les artefacts sont locaux et ignorés par Git ; les scripts et ce bilan sont versionnés. Aucun fichier Designer ou agencement WinForms n’a changé dans ce lot. Aucun pourcentage de couverture n’a été remesuré.

## Conditions nécessaires pour terminer les qualifications restantes

1. Fichier SOLIDWORKS `.swp` jetable identifié et autorisé pour modification/sauvegarde, puis essai réel de rechargement et de signature.
2. Contrôles ActiveX tiers autorisés par l’hôte pour tester leurs setters/événements ; les refus actuels sont des résultats de qualification, pas des succès d’hébergement.
3. Profils natifs VBE de langues supplémentaires et écrans/DPI différents, puis ancrage du chat vérifié dans ces profils.
4. Matrice de propriétés/événements supplémentaire et arbres COM/espions représentatifs. Les gardes qui empêchent les mutations précédemment corrélées à des crashs restent des fonctions limitées.

## Compilation et suite locale finale

Compilation : **0 erreur, 0 avertissement**. Suite Unit complète : **909 réussis, 0 échec, 0 ignoré**, rapport `artifacts/native-qualification/tests/unit-final.trx`. Ce passage Unit ne remplace pas les essais hôtes absents. Les 1 000 valeurs du rapport du débogueur ont aussi été contrôlées individuellement contre la fixture `index * 3`. Enregistrement COM, AccessVBOM et nettoyage du scénario final vérifiés restaurés (`registered-final/cleanup.json`), aucune instance Excel/SOLIDWORKS laissée ouverte.

## Lot IDE et intégration PR #4

Le test `ExcelParameterRenameTests.PrivateParameterRenamePreservesNativeCallBindingAndIsUndoable` utilise une nouvelle instance Excel visible et un classeur jetable possédé par la fixture. Il crée un module standard, prévisualise le renommage sans écriture, renomme le paramètre d'une Function privée et ses appels nommés, puis lance une Sub publique avec un argument nommé en omettant un paramètre optionnel. Les cellules A1 et B1 sont relues directement par COM Excel : **42** et **named call**. L'annulation du renommage restitue la source exacte.

Le même scénario enregistre un `.xlsm` et vérifie une copie isolée du fichier sauvegardé. Excel garde le document ouvert pour écriture ; le vérificateur refuse ce verrou concurrent et conserve une lecture immuable pendant le calcul de SHA et WinVerifyTrust. Avec le SIP VBA dédié absent, la réponse native est **VerifierUnavailable**, jamais un digest valide. La vérification cryptographique de signatures valides/altérées reste **NOT_RUN**, en attente de l'autorisation d'enregistrer temporairement le SIP Microsoft. La matrice locale couvre les résultats HRESULT et le nettoyage natif, sans être présentée comme une preuve cryptographique dans Office.

Preuve du lot copie/signature : `artifacts/pr4-integration/signature-snapshot/matrix.trx`, **4 tests réussis**, dont le scénario Excel. Le lot WinForms de la PR #4 passe **27 DesignSurface** (chargement, redimensionnement, propriétés et composants enfants), ainsi que `Test-ChatDesigner`. Les nouvelles vues ont leurs tests de cycle de vie et leurs fichiers Designer ; aucun essai SOLIDWORKS n'est ajouté à cette qualification.

## Lot complémentaire classes, ParamArray et options

Le lot du 28 septembre ajoute trois scénarios Excel dans le projet VSTest. Dans `artifacts/vbe-next/native-fixed/native.trx`, les scénarios de renommage de membres privés de classe et de transport ParamArray réussissent. Le premier vérifie les résultats de calcul, un compteur d'invocations et le retour à la source originale après annulation des trois renommages (Function, Property Get/Let et Property Get/Set). Les appels de la fixture sont directs ; cet essai ne qualifie pas toutes les formes de références `Me`. Le second vérifie le ParamArray vide, les valeurs mixtes, trente arguments, un préfixe ByVal et un tableau conservé comme argument unique.

La persistance des polices et palettes par catégorie reste **À qualifier**. Un premier essai a révélé une notification de sélection manquante. Après correction, la fermeture concurrente de l'instance Excel par une autre session a interrompu le scénario et sa restauration : `artifacts/vbe-next/native-final/native.trx`. Ce passage échoué n'est pas une preuve de persistance. La fixture conserve désormais son instantané complet en pièce jointe avant mutation et les deux erreurs lorsqu'une restauration échoue aussi. La reprise et la restauration des préférences attendent un créneau Excel coordonné ; aucun succès de restauration n'est revendiqué pour cette interruption.

Les palettes dessinées par le VBE sont exposées par indices natifs observés (`NativeIndex:n`), sans inventer des noms ni des valeurs RGB. Une taille de police lisible avec un catalogue de choix vide reste non modifiable par ce contrat. Les contrôles Win32 détenus des tests locaux vérifient le protocole de sélection ; ils ne remplacent pas la qualification des couleurs dans le VBE réel.
