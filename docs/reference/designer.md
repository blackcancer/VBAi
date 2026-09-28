# Concepteur de formulaires VBE

État du code après `2197c43`, le 28 septembre 2026. Ce guide décrit les UserForms VBA/MSForms ; l’édition des fenêtres de l’add-in dans Visual Studio est décrite dans [Concepteurs WinForms](../winforms-designer.md).

## Lire avant de modifier

`form_tree` expose l’arbre des contrôles et conteneurs avec leurs chemins canoniques. Frames, Pages de MultiPage et Tabs de TabStrip ont des chemins distincts. Utiliser la version d’arbre courante et le chemin retourné, plutôt qu’un nom supposé unique.

Les commandes de propriétés utilisent les descripteurs réels et une conversion typée : chaînes, nombres, couleurs et énumérations. Les propriétés en lecture seule, objets sans membre modifiable et écritures protégées produisent une erreur explicite. `ReadOnly=False` dans un inventaire ne suffit pas pour autoriser toute valeur COM.

| Besoin | Outils représentatifs |
| --- | --- |
| Inspecter | `form_tree`, `list_form_control_types`, `form_control_properties` |
| Créer/modifier | `add_form_control`, `add_nested_form_control`, `set_form_property`, `set_form_node_property` |
| Images | `set_form_picture`, `set_form_node_picture` ; chemin local explicitement fourni |
| Disposition | Prévisualisation et application de plans, dimensions, alignement, espacement et ordre |
| Événements | Inventaire des événements et création de procédure ; SHA du code et version d’arbre |
| Copier/récupérer | Duplication bornée, presse-papiers et historique natifs, récupération de coupe conservée en mémoire |
| Exécuter | `run_form`, `form_run_status`, `read_runtime_forms` ; relire l’état et les diagnostics |

Les noms exacts et champs requis de toutes les commandes figurent dans [le catalogue LLM](vbe-tools.md).

## Listes ComboBox et ListBox

`set_form_list_initializer` génère uniquement un bloc marqué de `UserForm_Initialize` pour une liste non liée. Fournir **soit `Items`, soit `Rows`**, avec le SHA courant du module et la version de l’arbre. `Rows` est une matrice rectangulaire de 0 à 64 lignes et de 1 à 10 colonnes, avec cellules de 256 caractères au maximum, sans caractères de contrôle. Le nombre de colonnes doit correspondre au contrôle. Un tableau vide efface la liste au runtime.

La commande préserve le code utilisateur hors du bloc géré et refuse un bloc modifié. Elle ne peuple pas directement la liste du designer. Les liaisons de listes ont leurs commandes dédiées et dépendent des capacités de l’hôte ; ne pas annoncer une persistance d’`AddItem` en conception.

## Historique et récupération

La duplication COM, l’historique natif du concepteur et la récupération après coupe sont des mécanismes distincts. Une écriture COM ne rejoint pas nécessairement la pile Undo native. La récupération conserve les formats de presse-papiers lisibles dans la session et vérifie l’état du formulaire avant restauration ; elle ne prouve pas une fidélité universelle des formats binaires ou des contrôles tiers. Les gestionnaires d’événements ne sont pas transférés implicitement avec les contrôles.

Un statut de commande retournée n’est pas une preuve d’affichage, d’initialisation réussie ou de fermeture d’un formulaire modal. Vérifier les fenêtres runtime et les diagnostics dans une lecture distincte.

## Inventaires et preuves

- [Propriétés des contrôles Excel](excel-control-properties.csv) et [propriétés d’un UserForm](excel-userform-properties.csv) : captures historiques de types/descripteurs, pas un journal de chaque écriture validée.
- [Exploration des propriétés](../archive/exploration/forms-property-coverage.md) : essais, refus et limites par propriété.
- [Exploration des duplications](../archive/exploration/forms-duplication-coverage.md) : profils qualifiés et refus.
- [Journal VBE](../archive/exploration/vbe-remaining-coverage.md) : qualifications natives des dispositions, listes, historiques et presse-papiers.
- [Travaux restants](../roadmap.md) : qualification à compléter par type, valeur, persistance et hôte.
