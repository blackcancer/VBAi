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

## Orchestration Designer, clipboard et exécution

| Source | Matrice ajoutée avant exécution |
| --- | --- |
| VbeForms.History | Action/révisions/availability, différence native observée ou absente, exceptions d'exécution et de readback avec conservation de la première erreur. |
| VbeForms.Run | Champs requis, contexte absent, identité projet/formulaire, mode/code/arbre/focus/commande changés après queue, pending/running/completed/failed, requête capturée, rejet Post, rétention de vingt opérations. |
| VbeForms.Layout | Paths canoniques et même conteneur, Pages/Tabs refusés, preview sans mutation, application root/nested, quatre coordonnées et finitude du readback, rollback complet/incomplet, tab order exhaustif et unique. |
| VbeForms.Clipboard | Sélection root/nested et limite, séquence instable, révisions, copy/paste/cut natifs, cut sans données/effet ou partiel, conservation des huit recoveries, restore et identité live, setters de sélection ignorés ou refusés. |
| VbeForms.CutRecovery | Recovery absent/incomplet/consommé, identité/container, mutation après coupe, coupe native partielle, restore clipboard invalide, paste unavailable/exception, noms/géométrie/tab order rétablis ou refusés, readback échoué et première erreur préservée. |
| VbePropertyInfo | Chaque valeur nullable ReadOnly, override SetterStatus explicite et retour au statut calculé. |
| VbeSession.Execute (scénarios séparés) | Routage réel de onze commandes Forms ; sélection/coupe/restore/recovery sur le service de la session, layout preview/apply, tab order, undo, queue/status et binding VBA vérifiés par leurs effets. |

Les frontières clipboard sont typées et conservent GetClipboardSequenceNumber,
Clipboard.GetDataObject et Clipboard.SetDataObject comme défauts. La fixture
remplace seulement ces transports par des IDataObject en mémoire et restaure
les delegates et descripteurs en Dispose. Aucun clipboard utilisateur ni hôte
Excel/SOLIDWORKS réel n'est utilisé. La récupération est créée par une véritable
NativeClipboard(cut), jamais par fabrication d'un objet recovery.

La condition CutTree conserve toutes ses validations, avec le test afterTree
avant error pour exercer les deux résultats indépendants sans changement de
sémantique. Dans Layout, TreeContainsPath accepte uniquement les paths générés
par ReadTreeNode, commençant tous par Controls/. Le path local ne subit aucune
mutation/réaffectation avant LastIndexOf ; seul end<0 redondant est supprimé.
Refus Pages/Tabs, canonicalisation, révisions et autres gardes restent présents.

Validation cumulative finale : **227 tests verts, zéro ignoré**, build sans
warning/error. Les **31 fichiers Forms atteignent 100 % lignes et branches** :
3673/3673 lignes et 6566/6566 branches. Rapport :
`artifacts/coverage/forms-complete-final/7cb00d24-cfa8-4e29-a6ef-196cc011f8b5/coverage.cobertura.xml`.
Le collector exclut uniquement `[ProviderTests]*` ; aucune exclusion production.

| Source | Lignes | Branches |
| --- | --- | --- |
| DesignerClipboardBackup.cs | 41/41 | 36/36 |
| FormHistoryDiff.cs | 53/53 | 36/36 |
| FormLayoutPlan.cs | 74/74 | 162/162 |
| OlePictureLoader.cs | 19/19 | 14/14 |
| VbeComEvents.cs | 93/93 | 32/32 |
| VbeComPropertyAccessors.cs | 84/84 | 48/48 |
| VbeControlCatalog.cs | 36/36 | 18/18 |
| VbeForms.CheckBoxDuplication.cs | 82/82 | 224/224 |
| VbeForms.Clipboard.cs | 127/127 | 246/246 |
| VbeForms.ComboBoxDuplication.cs | 80/80 | 208/208 |
| VbeForms.CommandButtonDuplication.cs | 76/76 | 206/206 |
| VbeForms.cs | 1092/1092 | 1566/1566 |
| VbeForms.CutRecovery.cs | 54/54 | 106/106 |
| VbeForms.FrameCopyPlan.cs | 79/79 | 90/90 |
| VbeForms.FrameDuplication.cs | 79/79 | 226/226 |
| VbeForms.FrameLabelDuplication.cs | 137/137 | 410/410 |
| VbeForms.FrameSimpleDuplication.cs | 142/142 | 348/348 |
| VbeForms.History.cs | 23/23 | 42/42 |
| VbeForms.LabelDuplication.cs | 120/120 | 322/322 |
| VbeForms.Layout.cs | 73/73 | 208/208 |
| VbeForms.ListBinding.cs | 50/50 | 100/100 |
| VbeForms.ListInitializer.cs | 219/219 | 306/306 |
| VbeForms.ListItems.cs | 81/81 | 78/78 |
| VbeForms.ListMutation.cs | 101/101 | 186/186 |
| VbeForms.ListRemoval.cs | 67/67 | 130/130 |
| VbeForms.OptionButtonDuplication.cs | 76/76 | 206/206 |
| VbeForms.ProfiledFrameDuplication.cs | 263/263 | 466/466 |
| VbeForms.Run.cs | 78/78 | 110/110 |
| VbeForms.TextBoxDuplication.cs | 80/80 | 208/208 |
| VbeForms.ToggleButtonDuplication.cs | 80/80 | 222/222 |
| VbePropertyInfo.cs | 14/14 | 6/6 |
