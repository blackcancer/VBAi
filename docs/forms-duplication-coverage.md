# Duplication des contrôles du concepteur VBE

Cette matrice suit une capacité précise : créer une copie d'un contrôle existant dans son conteneur. `form_tree`, `set_form_node_property` et les opérations de collection couvrent d'autres aspects du concepteur ; une copie partielle ne signifie pas que les propriétés omises sont inaccessibles.

| Type natif | Copie | Propriétés copiées | Preuve Excel | Limite actuelle |
| --- | --- | --- | --- | --- |
| Label | Bridge uniquement, partielle | Name, Caption, Left, Top, Width, Height, BackColor, Font.Name, Font.Size, Font.Bold | Deux essais sur classeur jetable, PID 3724 et 48020 : 1→2 nœuds, version changée, Excel vivant | Autres propriétés, images et ordre Z non copiés |
| TextBox | Bridge uniquement, partielle | Name, Left, Top, Width, Height, Value **si texte ou vide** | `Test-TextBoxDuplication.ps1` sur Excel PID 35452 : 1→2 nœuds, version changée, Excel vivant | Autres propriétés, liaisons et ordre Z non copiés |
| CheckBox | Bridge uniquement, partielle | Name, Caption, Left, Top, Width, Height, Value **booléen uniquement** | `Test-CheckBoxDuplication.ps1` sur Excel PID 14672 : 1→2 nœuds, valeur booléenne copiée, Excel vivant | TriState/null, autres propriétés et ordre Z non copiés |
| ComboBox | Bridge uniquement, partielle | Name, Left, Top, Width, Height, ListWidth **textuel uniquement** | `Test-ComboBoxDuplication.ps1` sur Excel PID 44860 : 1→2 nœuds, ListWidth copiée, Excel vivant | Items, liaisons, sélection et autres propriétés non copiés |
| CommandButton | Bridge uniquement, partielle | Name, Caption, Left, Top, Width, Height | `Test-CommandButtonDuplication.ps1` sur Excel PID 37444 : 1→2 nœuds, six propriétés copiées, Excel vivant | Événements et autres propriétés non copiés |
| Frame | Bridge uniquement, partielle | Name, Caption, Left, Top, Width, Height **si Frame vide** | `Test-FrameDuplication.ps1` sur Excel PID 48456 : copie vide, puis refus avant mutation avec Label enfant ; TreeVersion inchangée, Excel vivant | Copie récursive et autres propriétés non testées |
| Image | Non qualifiée | — | Ajout et lecture des propriétés | Image et copie non testées |
| ListBox | Non qualifiée | — | Ajout et lecture des propriétés | Items et copie non testés |
| MultiPage | Non qualifiée | — | Pages et enfants ajoutés, lus et supprimés | Copie récursive non testée |
| OptionButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| ScrollBar | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| SpinButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| TabStrip | Non qualifiée | — | Tabs ajoutés, lus et supprimés | Copie des Tabs non testée |
| ToggleButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |

Les contrôles ActiveX installés hors des 14 types natifs sont des candidats à valider dans le VBE hôte. L'essai `MSComctlLib.ListViewCtrl.2` a été refusé par `Controls.Add` dans Excel ; il n'entre dans aucune promesse de copie.

Le premier clonage générique fondé sur `PropertyDescriptor.IsReadOnly` a échoué sur `Label.Cancel`, suivi d'un crash Excel corrélé dans le temps. `ITypeInfo` annonce pourtant `PROPERTYPUT` pour ce membre. La liste positive est donc déterminée par des essais de mutation et de relecture ciblés ; ni le drapeau du descripteur ni la déclaration typelib ne suffisent. Les commandes de copie restent hors des outils LLM tant que chaque type et ses limites ne sont pas validés en hôte.

`frame_copy_plan` prépare en lecture seule une éventuelle copie d'un Frame avec Labels directs. Elle relève les enfants réels via `Parent`, propose des noms et chemins dans le futur conteneur, et signale les collisions, types non pris en charge et noms trop longs. `EligibleForLimitedProbe` décrit seulement la forme des données ; `MutationVerified=false` reste explicite. Le plan doit être relu dans Excel avant toute commande d'écriture récursive.
