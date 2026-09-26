# Couverture du VBE — état exploratoire

État relevé le 26 septembre 2026 sur le VBE 64 bits d'Excel, avec le complément chargé dans un `Classeur1` jetable. Les commandes du complément passent par l'objet VBE du processus hôte. Les sondes de fenêtres utilisent Win32/UI Automation ; aucun raccourci clavier ni clic à coordonnées n'est requis.

| Surface | Accès démontré | Lacune avant une commande LLM complète |
| --- | --- | --- |
| Projets et composants | `list_projects`, `list_modules`, `read_module`, `replace_lines` avec empreinte du module ; mode du projet lisible. | Créer, renommer, supprimer, importer/exporter des composants ; références ; chemins et identité de l'hôte ; sauvegarde explicite. |
| Éditeur de code | Texte vivant et sélection d'une ligne par `CodePane.SetSelection`. | Navigation par procédure/symbole, recherche globale, changements structurés, lecture fiable de l'erreur de compilation. |
| Formulaires | `list_forms`, `create_form`, `open_form`, `form_state`, `form_properties`. Le concepteur est une fenêtre MDI native `DesignerWindow` avec un `ThunderDFrame` pour le formulaire. Le chat Codex a invoqué `create_form` dans Excel. | Modifier les propriétés propres au formulaire, inspecter et modifier toutes les propriétés des contrôles, gérer les événements et les conteneurs. |
| Contrôles de formulaire | Ajout d'un Label, renommage, Caption, géométrie et police validés par relecture COM dans Excel. `form_control_properties` énumère les propriétés typées des contrôles. Les modifications exigent la version lue du formulaire. | Écriture des propriétés au-delà du sous-ensemble actuel ; suppression, duplication, ordre de tabulation et superposition. |
| Fenêtre Propriétés | Fenêtre native visible (`wndclass_pbrs`), propriété du formulaire énumérable via VBIDE. | La sélection dans la fenêtre Propriétés peut rester sur `Feuil1` alors que le concepteur du formulaire est actif ; ne pas la prendre comme source de vérité implicite. |
| Explorateur de projets et d'objets | Fenêtre Projet native visible ; commande native de l'Explorateur d'objets découverte (`Id=473`, activée). | Navigation et lecture structurée de l'Explorateur d'objets non vérifiées. |
| Références | Commande native « Références... » découverte (`Id=942`, activée). Sur la branche expérimentale, `list_references` lit les références sélectionnées avec GUID, version, chemin et `IsBroken`. L'ajout puis le retrait de Scripting Runtime ont été relus dans Excel ; l'empreinte finale est revenue à la valeur initiale. | Résolution des références cassées et cas d'échec d'ajout/retrait non vérifiés ; les 651 entrées disponibles dans la boîte native ne sont pas toutes sélectionnées. |
| Débogage | Mode et sélection lus par VBIDE ; commandes natives inventoriées. Un cycle arrêt/pas détaillé/poursuite a été observé dans un essai Excel précédent. | Inventaire fiable des points d'arrêt, pile d'appels, exceptions et vérification systématique de l'effet après chaque commande. |
| Fenêtres Exécution et Variables locales | Fenêtres natives visibles (`VbaWindow`) ; lectures par UI Automation validées précédemment dans Excel. | Écriture/évaluation dans Exécution, structure typée des variables, espions et pile d'appels. |
| Menus et barres d'outils | `list_commands` lit chemins, intitulés, identifiants et état activé. Cette session révèle notamment `Id=222` Propriétés, `Id=1820` Ajouter un espion. | `list_commands` limite les résultats à 200 ; les identifiants/intitulés doivent être relevés dans chaque hôte et langue. Exécution générique avec contrôle d'effet à concevoir. |
| Fenêtres/options VBE | Arbre HWND du processus accessible en lecture seule, sans focus. Une fenêtre outil native a été créée et liée au bureau VBE dans Excel (voir ci-dessous). | Positionnement à droite, contenu du panneau, persistance de la disposition, boîtes modales et boîte à outils à vérifier. |

## Fenêtre outil ancrable du complément

Le 26 septembre 2026, dans Excel 64 bits (PID 31488), une sonde chargée dans le processus VBE a appelé `Microsoft.Vbe.Interop._Windows.CreateToolWindow` avec l'instance `Microsoft.Vbe.Interop.AddIn` obtenue de `VBE.Addins.Item("CodexVBE.AddIn")`, le ProgID `Shell.Explorer.2`, un titre et un GUID de fenêtre. Le dernier paramètre est un **`ref object`** dans l'interface Microsoft : l'appel dynamique initial avec `out object` échouait à la conversion de l'argument COM. L'appel typé a retourné une fenêtre `Type=15` et un objet document COM. Après `window.Visible = true` puis `VBE.MainWindow.LinkedWindows.Add(window)`, le compteur `MainWindow.LinkedWindows.Count` valait 4. L'arbre Win32 montrait un `GenericPane` visible « CodexVBE Dock Probe » sous la fenêtre principale VBE et un enfant `Shell Embedding`.

`Forms.Frame.1` a retourné `E_FAIL` et `Forms.UserForm.1` `CO_E_CLASSSTRING` dans ce même essai. `Shell.Explorer.2` prouve la création et l'ancrage d'une fenêtre outil, mais seulement avec le document ActiveX navigateur vide : le chat n'y est pas encore hébergé et le côté droit précis n'a pas été validé. La sonde d'essai a été retirée des sources du complément après cette vérification. Pour la réalisation, créer la fenêtre avec l'interface typée, conserver les objets `Window` et `DocObj`, puis vérifier le contenu et la disposition dans Excel et SOLIDWORKS.

## Session Excel observée

- Processus Excel 15012 ; complément connecté ; projet `VBAProject` en mode Conception (`2`).
- Formulaire `CodexFormProbe` ouvert. Son contrôle `lblEtat` affiche « Prêt pour Codex », position `(24, 24)`, dimensions `(120, 24)`, police Segoe UI 14 gras. La relecture `form_state` retourne la version `612b0bfa9c3c8735da385de2e0ea452c65f0ab15e75d89ad816bfdb07fd059bd`.
- Le formulaire n'est pas enregistré sur disque ; cette preuve concerne l'état vivant du VBE Excel. L'instance reste visible pour les essais.
- La fenêtre Propriétés affiche `Feuil1` tandis que le titre MDI actif indique `CodexFormProbe (UserForm)`. Les commandes doivent cibler explicitement projet, formulaire et contrôle, et relire l'objet COM ciblé.

## Inventaire des propriétés de contrôle

Un nouvel essai Excel visible (PID 40240) a placé les 14 contrôles standard sur `CodexAllControlProperties` et exporté **771 descripteurs** dans [excel-control-properties.csv](excel-control-properties.csv). La sonde reproductible est `tools/probes/Export-ControlProperties.ps1`. Chaque ligne contient le type de contrôle, son ProgID, le nom et le type de la propriété, `ReadOnly`, la valeur lisible et une éventuelle erreur. Les nombres de propriétés par type vont de 36 (SpinButton) à 86 (ComboBox). Les onze erreurs de lecture portent toutes sur `_Font_Reserved` (`0x80020003`). Cet export est un inventaire de lecture sur cet hôte Excel ; `ReadOnly=False` n'établit pas encore qu'une écriture soit acceptée par le VBE. La boîte à outils peut aussi accueillir des contrôles ActiveX supplémentaires, qui devront être découverts et inspectés dynamiquement.

Dans des classeurs Excel jetables, `TypeDescriptor.GetProperties` sur les objets COM du concepteur a retourné les résultats suivants. Les 14 ProgID sont acceptés par `add_form_control` et ont été relus via `form_control_properties` dans le VBE Excel vivant. « Modifiables » signifie seulement que le descripteur ne porte pas `IsReadOnly` ; cela ne prouve pas que toute écriture soit sûre.

| Contrôle `Forms.*.1` | Propriétés | Descripteurs modifiables | `Caption` | `Value` | `Picture` |
| --- | ---: | ---: | :---: | :---: | :---: |
| Label | 51 | Non relevé | Oui | Non relevé | Non relevé |
| TextBox | 79 | Non relevé | Non relevé | Non relevé | Non relevé |
| CommandButton | 49 | Non relevé | Oui | Non relevé | Non relevé |
| CheckBox | 57 | Non relevé | Oui | Non relevé | Non relevé |
| Frame | 63 | Non relevé | Oui | Non relevé | Non relevé |
| ComboBox | 86 | Non relevé | Non relevé | Non relevé | Non relevé |
| OptionButton | 57 | 48 | Oui | Oui | Oui |
| ListBox | 64 | 54 | Non | Oui | Non |
| SpinButton | 36 | 29 | Non | Oui | Non |
| ScrollBar | 38 | 31 | Non | Oui | Non |
| Image | 38 | 31 | Non | Non | Oui |
| MultiPage | 45 | 36 | Non | Oui | Non |
| TabStrip | 51 | 38 | Non | Oui | Non |
| ToggleButton | 57 | 48 | Oui | Oui | Oui |

Les noms et types incluent `Name`, `Left`, `Top`, `Width`, `Height`, `Visible`, `Enabled`, puis des propriétés propres à chaque contrôle. `_Font_Reserved` échoue à la lecture (`0x80020003`) sur Label, TextBox, CommandButton, CheckBox, Frame, ComboBox, OptionButton, ListBox, MultiPage, TabStrip et ToggleButton ; aucune autre erreur de lecture n'a été relevée sur les huit derniers types testés. Le ComboBox a nécessité une garde pour un descripteur dont le type est nul.

### Propriétés propres au UserForm

Dans le classeur Excel jetable suivant, le chat a créé `CodexAgentFormProbe`. `form_properties` a relu 50 propriétés du `VBComponent.Properties`, parmi lesquelles `Name`, `Caption=UserForm1`, `Left=0`, `Top=0`, `Width=240`, `Height=180`, `Enabled=True`, `Tag`, `BackColor`, `ForeColor`, `StartUpPosition=1`, `ShowModal=True` et `Zoom=100`. Certaines sont des objets ou états non scalaires (`Controls`, `Font`, `Selected`, `ActiveControl`) ; la commande de lecture actuelle les rend `null` et ne donne ni type ni caractère modifiable. Aucune commande exposée ne change encore les propriétés du formulaire lui-même. La création du composant n'établit donc pas une capacité de conception complète du UserForm. Une commande de modification doit cibler la propriété, convertir la valeur selon son type, rejeter les objets/états non éditables, puis relire la propriété et le formulaire. La révision actuelle ne couvre que `Caption`, `Width`, `Height` et les contrôles du premier niveau : toute autre propriété écrite nécessite aussi une protection de concurrence adaptée.

### Conteneurs et contrôles imbriqués

`MultiPage.Pages` et `TabStrip.Tabs` sont exposés comme collections COM par leurs objets du concepteur ; `form_control_properties` les présente actuellement avec une valeur `null`, car il ne sérialise que les scalaires. La [référence Microsoft Forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/objects-microsoft-forms) distingue trois collections : `Controls` sur UserForm, Frame ou Page, `Pages` sur MultiPage et `Tabs` sur TabStrip. L'[exemple Microsoft](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/page-object-multipage-control-add-clear-remove-methods-example) montre l'ajout d'un contrôle via `MultiPage1.Pages(0).Controls.Add(...)`. Les index de `Controls` commencent à zéro d'après la [référence de la collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms).

Les commandes de mutation actuelles résolvent seulement `form.Designer.Controls` et donc les contrôles du premier niveau. Il faut des chemins typés (`Frame`, `MultiPage/Page`, `TabStrip/Tab`) pour créer, modifier et supprimer leurs enfants. La révision exigée par ces commandes couvre les propriétés du formulaire et les contrôles du premier niveau ; les futures mutations imbriquées devront utiliser une empreinte récursive sans ambiguïté. Un essai COM externe via `Excel.Application.VBE` a retourné une collection de projets vide dans cette session, alors que le pont exécuté **dans** le complément voit le projet `VBAProject`.

Sur la branche expérimentale `feat/llm-commands` chargée dans Excel le même jour, `form_tree` a ensuite prouvé la **lecture récursive** : un formulaire avec `fraProbe` (Frame), `mpgProbe` (MultiPage) et `tabsProbe` (TabStrip) retourne sept nœuds. Les chemins des enfants sont `Controls/mpgProbe/Pages/Page1`, `.../Page2`, `Controls/tabsProbe/Tabs/Tab1` et `.../Tab2`. Les pages ont 38 propriétés chacune et les onglets huit. La mutation des pages, onglets et enfants n'est pas encore exposée au LLM.

Une sonde ultérieure **a ajouté** `lblInsideFrame` dans `Frame.Controls` et `lblInsidePage` dans `MultiPage.Pages("Page1").Controls` sur un autre formulaire Excel jetable. `form_tree` a relu les deux chemins imbriqués. Elle a aussi révélé que `form.Designer.Controls` énumère ces deux contrôles en plus des contrôles du premier niveau : la première version de `form_tree` les compte à tort une seconde fois comme `Controls/lblInsideFrame` et `Controls/lblInsidePage` (huit nœuds au lieu de six objets uniques). Le parcours de racine doit filtrer selon le `Parent` réel avant d'exposer des mutations hiérarchiques. Cette sonde n'est pas encore une validation des propriétés modifiables de chaque nœud.

La même branche a permis de modifier et relire `Caption`, `Width`, `Height`, `ShowModal`, `StartUpPosition`, `Tag` et `Font.Name` sur un formulaire jetable, avec changement de sa révision. `Picture` a été affecté à partir d'un BMP local de 2×2 pixels : l'empreinte `stdole.IPictureDisp` et la révision sont restées stables sur deux relectures successives sans mutation. Ce test ne démontre pas encore tous les formats d'image ni toutes les autres propriétés. Une première implémentation de `Font.Name` utilisant `Property.Object` avait donné un faux succès, puis Excel a planté (`0xc0000374`) ; elle a été retirée. La voie corrigée utilise l'objet du concepteur et a passé un essai ciblé.

### Références du projet

Dans le même hôte Excel, la boîte `Outils > Références...` est une fenêtre `#32770` avec une `ListBox` native dessinée par le VBE. Les messages de lecture Win32 donnent 651 noms **disponibles**, mais UI Automation n'expose aucun enfant de cette liste et ne donne donc pas l'état coché. La commande `list_references` exécutée dans le complément lit directement `VBProject.References` : le projet `VBAProject` jetable a quatre références sélectionnées (`VBA` 4.2, `Excel` 1.9, `stdole` 2.0, `Office` 2.8), toutes non cassées, avec GUID et chemin absolu. C'est la source à utiliser pour les commandes futures d'ajout, de retrait et de résolution.

Dans un autre `Classeur1` jetable, `add_reference_guid` a ajouté `{420B2830-E718-11CF-893D-00A0C9054228}` avec version demandée `0.0` : VBIDE a sélectionné **Scripting Runtime 1.0** depuis `C:\Windows\System32\scrrun.dll`, faisant passer la liste de quatre à cinq références. `remove_reference` a ciblé ce GUID et sa version exacte avec l'empreinte de liste lue après l'ajout ; la liste est revenue à quatre références et son SHA-256 est redevenu identique à celui du départ. Ce cycle ne valide pas encore `add_reference_file` ni les bibliothèques cassées.

Un essai ultérieur dans Excel PID 40240 a validé `add_reference_file` sur le même `scrrun.dll` : cinq références avant, six après l'ajout, puis cinq après `remove_reference`. Le GUID, le nom `Scripting` et le chemin ont été relus ; l'empreinte finale de la collection est égale à l'empreinte initiale. La cinquième référence initiale est due au formulaire MSForms du classeur de test. Les références cassées restent à explorer.

## Prochaine exploration

1. Lire les collections `Frame.Controls`, `MultiPage.Pages`, `Page.Controls` et `TabStrip.Tabs` depuis le complément dans Excel ; vérifier ajout, renommage, géométrie et suppression sur un classeur jetable, puis étendre la version récursive.
2. Vérifier les références, les procédures, la boîte à outils et les événements via les API VBIDE/MSForms dans Excel.
3. Tester les fenêtres de débogage restantes et relire l'effet des commandes natives, avant toute promesse de couverture complète.
4. Rejouer les commandes retenues dans SOLIDWORKS seulement après ouverture de son VBE par l'utilisateur ; ne pas déduire sa compatibilité du seul essai Excel.
