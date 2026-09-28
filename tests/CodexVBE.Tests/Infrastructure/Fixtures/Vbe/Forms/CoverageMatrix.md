# Matrice Forms après fusion chat-ux

## Plans, sauvegarde et liaison de listes

Les scénarios sont ajoutés par matrice avant la première exécution du fichier.
Les sauvegardes utilisent seulement des `IDataObject` en mémoire. Les contrats
Designer existants exercent les vrais chemins de validation et génération VBA.

| Source | Matrice |
| --- | --- |
| DesignerClipboardBackup | Source absente, format inconnu, bag absent/vide/non binaire, copie indépendante des octets, formats null omis, chaînes, taille cumulée maximale, concordance et différences de formats. |
| FormHistoryDiff | Collections absentes/null/malformées, propriétés ajoutées/modifiées/illisibles, contrôles imbriqués supprimés/ajoutés, type de contrôle et flags CanUndo/CanRedo/CanPaste exclus de la comparaison d'éditions. |
| FormLayoutPlan | Toutes les actions sur les deux axes, bornes des counts et conteneurs, finitude et tailles d'entrée, spacing/distribution avec chevauchement refusé, alignements/sizing/centrage/grid, débordements Inf/NaN et coordonnées hors conteneur. |
| VbeForms.ListBinding | Révisions et chemin hôte requis, workbook absent/relatif/différent/non Excel, arbre périmé, contrôle absent ou autre type, code périmé, ComboBox/ListBox, noms littéraux et bornes Excel indépendantes, remplacement du bloc géré sans écriture Designer. |
| VbeForms.ListInitializer | Budget de ligne VBA, tous segments et collections de chemin, identifiants invalides, en plus de la matrice existante de création/remplacement/idempotence/erreurs et contenu généré. |

Simplification validée de FormLayoutPlan : les Width/Height sont initialement
finies et positives ; le résultat en copie les valeurs, et same_width,
same_height et same_size utilisent uniquement l'ancre validée. Aucune autre
réaffectation ni mutation externe ne peut rendre ces tailles négatives/nulles.
Les deux comparaisons finales Width<=0/Height<=0 sont retirées ; validation
initiale, finitude finale et limites du conteneur sont conservées et testées.

Validation de ce lot : **45 tests verts, zéro ignoré** ; cinq fichiers à
**100 % lignes et branches**. Rapport du worktree :
`artifacts/coverage/forms-binding-validated/6e1441a6-f228-4a6b-a381-9a788335f7d1/coverage.cobertura.xml`.
Seule exclusion collector : `[ProviderTests]*`, aucune exclusion de production.
