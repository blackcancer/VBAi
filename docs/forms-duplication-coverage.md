# Duplication des contrôles du concepteur VBE

Cette matrice suit une capacité précise : créer une copie d'un contrôle existant dans son conteneur. `form_tree`, `set_form_node_property` et les opérations de collection couvrent d'autres aspects du concepteur ; une copie partielle ne signifie pas que les propriétés omises sont inaccessibles.

| Type natif | Copie | Propriétés copiées | Preuve Excel | Limite actuelle |
| --- | --- | --- | --- | --- |
| Label | Bridge uniquement, partielle | Name, Caption, Left, Top, Width, Height, BackColor, Font.Name, Font.Size, Font.Bold | Deux essais sur classeur jetable, PID 3724 et 48020 : 1→2 nœuds, version changée, Excel vivant | Autres propriétés, images et ordre Z non copiés |
| TextBox | Bridge uniquement, partielle | Name, Left, Top, Width, Height, Value **si texte ou vide** | `Test-TextBoxDuplication.ps1` sur Excel PID 35452 : 1→2 nœuds, version changée, Excel vivant | Autres propriétés, liaisons et ordre Z non copiés |
| CheckBox | Bridge uniquement, partielle ; essai en attente | Name, Caption, Left, Top, Width, Height, Value **booléen uniquement** | Value=true testé séparément ; duplication à tester avec `Test-CheckBoxDuplication.ps1` | TriState/null, autres propriétés et ordre Z non copiés |
| ComboBox | Non qualifiée | — | ListWidth testé séparément | Items, liaisons et copie non testés |
| CommandButton | Non qualifiée | — | Ajout et procédure Click testés | Copie non testée |
| Frame | Non qualifiée | — | Contrôles enfants ajoutés et lus | Copie récursive non testée |
| Image | Non qualifiée | — | Ajout et lecture des propriétés | Image et copie non testées |
| ListBox | Non qualifiée | — | Ajout et lecture des propriétés | Items et copie non testés |
| MultiPage | Non qualifiée | — | Pages et enfants ajoutés, lus et supprimés | Copie récursive non testée |
| OptionButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| ScrollBar | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| SpinButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| TabStrip | Non qualifiée | — | Tabs ajoutés, lus et supprimés | Copie des Tabs non testée |
| ToggleButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |

Les contrôles ActiveX installés hors des 14 types natifs sont des candidats à valider dans le VBE hôte. L'essai `MSComctlLib.ListViewCtrl.2` a été refusé par `Controls.Add` dans Excel ; il n'entre dans aucune promesse de copie.

Le premier clonage générique fondé sur `PropertyDescriptor.IsReadOnly` a échoué sur `Label.Cancel`, suivi d'un crash Excel corrélé dans le temps. `ITypeInfo` annonce pourtant `PROPERTYPUT` pour ce membre. La liste positive est donc déterminée par des essais de mutation et de relecture ciblés ; ni le drapeau du descripteur ni la déclaration typelib ne suffisent. Les deux commandes de copie restent hors des outils LLM tant que chaque type et ses limites ne sont pas validés en hôte.
