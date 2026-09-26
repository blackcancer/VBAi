# Couverture du VBE — état exploratoire

État relevé le 26 septembre 2026 sur le VBE 64 bits d'Excel, avec le complément chargé dans un `Classeur1` jetable. Les commandes du complément passent par l'objet VBE du processus hôte. Les sondes de fenêtres utilisent Win32/UI Automation ; aucun raccourci clavier ni clic à coordonnées n'est requis.

| Surface | Accès démontré | Lacune avant une commande LLM complète |
| --- | --- | --- |
| Projets et composants | `list_projects`, `list_modules`, `read_module`, `replace_lines` avec empreinte du module ; mode du projet lisible. | Créer, renommer, supprimer, importer/exporter des composants ; références ; chemins et identité de l'hôte ; sauvegarde explicite. |
| Éditeur de code | Texte vivant et sélection d'une ligne par `CodePane.SetSelection`. | Navigation par procédure/symbole, recherche globale, changements structurés, lecture fiable de l'erreur de compilation. |
| Formulaires | `list_forms`, `create_form`, `open_form`, `form_state`, `form_properties`. Le concepteur est une fenêtre MDI native `DesignerWindow` avec un `ThunderDFrame` pour le formulaire. | Définir les propriétés du formulaire, inspecter et modifier toutes les propriétés des contrôles, gérer les événements et les conteneurs. |
| Contrôles de formulaire | Ajout d'un Label, renommage, Caption, géométrie et police validés par relecture COM dans Excel. `form_control_properties` énumère les propriétés typées des contrôles. Les modifications exigent la version lue du formulaire. | Écriture des propriétés au-delà du sous-ensemble actuel ; suppression, duplication, ordre de tabulation et superposition. |
| Fenêtre Propriétés | Fenêtre native visible (`wndclass_pbrs`), propriété du formulaire énumérable via VBIDE. | La sélection dans la fenêtre Propriétés peut rester sur `Feuil1` alors que le concepteur du formulaire est actif ; ne pas la prendre comme source de vérité implicite. |
| Explorateur de projets et d'objets | Fenêtre Projet native visible ; commande native de l'Explorateur d'objets découverte (`Id=473`, activée). | Navigation et lecture structurée de l'Explorateur d'objets non vérifiées. |
| Références | Commande native « Références... » découverte (`Id=942`, activée). | Liste, ajout/retrait et résolution des références non vérifiés. |
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

### Conteneurs et contrôles imbriqués

`MultiPage.Pages` et `TabStrip.Tabs` sont exposés comme collections COM par leurs objets du concepteur ; `form_control_properties` les présente actuellement avec une valeur `null`, car il ne sérialise que les scalaires. La [référence Microsoft Forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/objects-microsoft-forms) distingue trois collections : `Controls` sur UserForm, Frame ou Page, `Pages` sur MultiPage et `Tabs` sur TabStrip. L'[exemple Microsoft](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/page-object-multipage-control-add-clear-remove-methods-example) montre l'ajout d'un contrôle via `MultiPage1.Pages(0).Controls.Add(...)`. Les index de `Controls` commencent à zéro d'après la [référence de la collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms).

Les commandes actuelles résolvent seulement `form.Designer.Controls` et donc les contrôles du premier niveau. Il faut des chemins typés (`Frame`, `MultiPage/Page`, `TabStrip/Tab`) pour parcourir, créer, modifier et supprimer leurs enfants. `VbeForms.Version` ne hache que les contrôles du premier niveau et quelques propriétés du formulaire ; une édition de page ou d'enfant ne changerait pas cette empreinte. Étendre l'empreinte récursivement avant de proposer des mutations imbriquées avec `ExpectedFormVersion`. Un essai COM externe via `Excel.Application.VBE` a retourné une collection de projets vide dans cette session, alors que le pont exécuté **dans** le complément voit le projet `VBAProject` ; les contrôles imbriqués n'ont donc pas encore été lus ou modifiés en direct.

## Prochaine exploration

1. Lire les collections `Frame.Controls`, `MultiPage.Pages`, `Page.Controls` et `TabStrip.Tabs` depuis le complément dans Excel ; vérifier ajout, renommage, géométrie et suppression sur un classeur jetable, puis étendre la version récursive.
2. Vérifier les références, les procédures, la boîte à outils et les événements via les API VBIDE/MSForms dans Excel.
3. Tester les fenêtres de débogage restantes et relire l'effet des commandes natives, avant toute promesse de couverture complète.
4. Rejouer les commandes retenues dans SOLIDWORKS seulement après ouverture de son VBE par l'utilisateur ; ne pas déduire sa compatibilité du seul essai Excel.
