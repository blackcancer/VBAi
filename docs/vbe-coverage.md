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
| Fenêtres/options VBE | Arbre HWND du processus accessible en lecture seule, sans focus. | Préférences, disposition, fenêtres ancrées, boîtes modales et boîte à outils à inventorier ; API publique hétérogène. |

## Session Excel observée

- Processus Excel 15012 ; complément connecté ; projet `VBAProject` en mode Conception (`2`).
- Formulaire `CodexFormProbe` ouvert. Son contrôle `lblEtat` affiche « Prêt pour Codex », position `(24, 24)`, dimensions `(120, 24)`, police Segoe UI 14 gras. La relecture `form_state` retourne la version `612b0bfa9c3c8735da385de2e0ea452c65f0ab15e75d89ad816bfdb07fd059bd`.
- Le formulaire n'est pas enregistré sur disque ; cette preuve concerne l'état vivant du VBE Excel. L'instance reste visible pour les essais.
- La fenêtre Propriétés affiche `Feuil1` tandis que le titre MDI actif indique `CodexFormProbe (UserForm)`. Les commandes doivent cibler explicitement projet, formulaire et contrôle, et relire l'objet COM ciblé.

## Inventaire des propriétés de contrôle

Dans des classeurs Excel jetables, `TypeDescriptor.GetProperties` sur les objets COM du concepteur a retourné 51 propriétés pour `Label`, 79 pour `TextBox`, 49 pour `CommandButton`, 57 pour `CheckBox`, 63 pour `Frame` et 86 pour `ComboBox`. Les noms et types comprennent notamment `Name`, `Left`, `Top`, `Width`, `Height`, `Visible`, `Enabled`, ainsi que `Caption`, `FontName`, `FontSize`, `FontBold`, `TabIndex` ou `BackColor` selon le type. Chaque propriété renvoie son état en lecture seule et une éventuelle erreur de lecture. Sur ces six types, `_Font_Reserved` est la seule propriété dont la lecture a échoué (`0x80020003`) ; les autres valeurs ne prouvent pas encore qu'une écriture soit sûre. Le ComboBox a nécessité une garde pour un descripteur dont le type est nul. Les autres contrôles MSForms restent à tester.

## Prochaine exploration

1. Inventorier les propriétés de chaque type de contrôle directement sur l'objet `Designer.Controls`, avec type, capacité de lecture/écriture et état de l'objet, sans supposer que tous offrent `Caption` ou `Font`.
2. Vérifier les références, les procédures, la boîte à outils et les événements via les API VBIDE/MSForms dans Excel.
3. Tester les fenêtres de débogage restantes et relire l'effet des commandes natives, avant toute promesse de couverture complète.
4. Rejouer les commandes retenues dans SOLIDWORKS seulement après ouverture de son VBE par l'utilisateur ; ne pas déduire sa compatibilité du seul essai Excel.
