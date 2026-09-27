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
| OptionButton | Bridge uniquement, partielle ; essai en attente | Name, Caption, Left, Top, Width, Height | Ajout et lecture des propriétés ; duplication à tester avec `Test-OptionButtonDuplication.ps1` | Value, GroupName, événements et autres propriétés non copiés pour préserver la sélection du groupe |
| ScrollBar | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| SpinButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |
| TabStrip | Non qualifiée | — | Tabs ajoutés, lus et supprimés | Copie des Tabs non testée |
| ToggleButton | Non qualifiée | — | Ajout et lecture des propriétés | Copie non testée |

Les contrôles ActiveX installés hors des 14 types natifs sont des candidats à valider dans le VBE hôte. L'essai `MSComctlLib.ListViewCtrl.2` a été refusé par `Controls.Add` dans Excel ; il n'entre dans aucune promesse de copie.

Le premier clonage générique fondé sur `PropertyDescriptor.IsReadOnly` a échoué sur `Label.Cancel`, suivi d'un crash Excel corrélé dans le temps. `ITypeInfo` annonce pourtant `PROPERTYPUT` pour ce membre. La liste positive est donc déterminée par des essais de mutation et de relecture ciblés ; ni le drapeau du descripteur ni la déclaration typelib ne suffisent. Les commandes de copie restent hors des outils LLM tant que chaque type et ses limites ne sont pas validés en hôte.

`frame_copy_plan` prépare en lecture seule une éventuelle copie d'un Frame avec Labels directs. Elle relève les enfants réels via `Parent`, propose des noms et chemins dans le futur conteneur, et signale les collisions, types non pris en charge et noms trop longs. Dans Excel PID 27988, elle a proposé le chemin canonique du Label et signalé le TextBox supplémentaire comme non pris en charge, sans changer `TreeVersion`. `EligibleForLimitedProbe` décrit seulement la forme des données ; `MutationVerified=false` reste explicite.

`duplicate_form_frame_labels` est une sonde bridge transactionnelle. Elle exige un Frame **racine** avec au moins un Label direct et aucun autre enfant. Elle précharge les champs des Labels selon le profil positif déjà testé, crée le Frame puis ses Labels, relit chaque propriété et l'arbre. Dans Excel PID 13396, deux Labels ont été copiés avec `NodeCount` 3→6 et chemins canoniques relus ; l'ajout ultérieur d'un TextBox enfant a fait refuser une seconde copie avant mutation. Si une étape d'écriture échoue, la commande tente de retirer les Labels créés en ordre inverse puis le Frame et signale une restauration incomplète. Les autres descendants, propriétés et événements ne sont pas copiés.

`frame_simple_copy_plan` et `duplicate_form_frame_simple_children` constituent une nouvelle sonde, encore à tester en Excel, pour un Frame racine contenant uniquement des Labels et TextBoxes directs. Les profils positifs de chaque enfant restent limités aux champs validés séparément : Label (libellé/géométrie/couleur/police) et TextBox (géométrie/Value textuel). Le plan historique `frame_copy_plan` et la commande Label seule conservent leur refus du TextBox ; la nouvelle route accepte ce cas mais refuse les autres types, les descendants supplémentaires et les liaisons de données. La mutation relit chaque enfant et prévoit un rollback explicite.
