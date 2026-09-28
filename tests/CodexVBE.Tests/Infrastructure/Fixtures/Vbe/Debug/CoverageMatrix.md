# Matrice de couverture Debug après fusion chat-ux

Les tests suivent les vrais chemins de navigation, lecture, validation et invocation.
Les contrats VBIDE construits dans `EditorDebugFixture` représentent des projets,
modules, fenêtres, commandes et volets dont les changements suivent les événements
natifs. Les delegates remplacent uniquement les appels user32 et l'acquisition UIA.
Les `AutomationElement`, patterns et recherches de l'Explorateur d'objets utilisent
le transport UIA réel vers des fournisseurs jetables avec `Application.Run` sur
leur propre thread STA. Aucun handle ne désigne Excel ou SOLIDWORKS.

| Fichier | Scénarios préparés avant la première exécution |
| --- | --- |
| VbeDebugWindows.CodeView | Identification exacte, ambiguïtés fenêtre/scrollbar/toolbar, visibilité et lecture des bounds, ascendants modaux, bornes de géométrie, clic local procédure/module. |
| VbeDebugWindows.RuntimeForms | Filtrage processus/classe/visibilité, propriétaire absent/désactivé, bounds absents/erreur, énumération échouée, comptage et troncature à 64. |
| VbeDebug.Editor | Action/caption/mode, module actif, topologie 0/1/2/3, entrées détruites et étrangères, commande exacte désactivée/absente, split/unsplit, aucune action si déjà conforme, erreurs exécution/relecture. |
| VbeDebug.CodePanes | Inventaire et expiration des tokens, module vide, volet détruit, modes et appartenance, fermeture/remplacement, versions de chaque volet, défilement et erreurs séparées, sélection/hash modifiés, focus et viewport, vérification des deux volets après toolbar native. |
| VbeDebug.NativeHistory | Projet unique, mode design, refus UserForm, hash stable et modules triés, commandes Undo/Redo exactes, fenêtre modale, version périmée, code inchangé/modifié, modules retirés/ajoutés, priorités des erreurs et vérification partielle. |
| VbeDebug.Arrangement | Identités code/Designer/browser, chemins absents/indisponibles, fenêtres filtrées, identités absentes/ambiguës, 0/1/2/65 fenêtres, version et commande exacte, trois orientations, géométrie invalide, erreurs exécution/relecture, topologie changée. |
| VbeDebug.NativeNavigation | Actions et dispatcher, contexte absent, sélection exacte, Request réutilisée après scheduling, Queued/Running/Observing/Completed/Failed, commandes/revision/mode/position changés avant invocation, focus, Post rejeté, erreurs observation, délai expiré, projet non enregistré, browser/pane fermé et rétention de 20 opérations. |
| VbeDebugWindows.ObjectBrowser | Lecture patterns et descriptions tronquées, accès UIA échoué, arguments et pagination, listes vides/ambiguës et langues FR/EN, sélection native exacte, notifications refusées, état relu différent, bibliothèque exacte et contrôles indisponibles, attente asynchrone de membres et délai épuisé. |

Toutes les frontières statiques sont restaurées lors du nettoyage, même en cas
d'échec. Le contexte de synchronisation est également restauré. Le délai de
navigation de deux secondes est testé avec le vrai temps ; les pauses user32/UIA
utilisent le delegate existant `PauseNative`, dont le défaut reste `Thread.Sleep`.

Une seule condition redondante a été retirée, après revue : le second
`handle == IntPtr.Zero` de `SelectBrowserLibrary`. La première validation refuse
zéro et la variable locale n'est ni réassignée ni passée par référence avant la
notification. La première validation, le contrôle de classe, les ascendants,
le pattern, `PostMessage` et la vérification de sélection restent testés.

Validation cumulée : **180 tests verts, zéro ignoré**. Collector : unique
exclusion `[ProviderTests]*`, aucune exclusion de production.

| Source | Lignes | Branches |
| --- | ---: | ---: |
| VbeDebug.Arrangement.cs | 104/104 | 210/210 |
| VbeDebug.CodePanes.cs | 119/119 | 256/256 |
| VbeDebug.cs | 589/589 | 1016/1016 |
| VbeDebug.Editor.cs | 33/33 | 78/78 |
| VbeDebug.NativeHistory.cs | 69/69 | 98/98 |
| VbeDebug.NativeNavigation.cs | 113/113 | 348/348 |
| VbeDebugWindows.CodeView.cs | 32/32 | 38/38 |
| VbeDebugWindows.cs | 1141/1141 | 738/738 |
| VbeDebugWindows.ObjectBrowser.cs | 199/199 | 198/198 |
| VbeDebugWindows.RuntimeForms.cs | 31/31 | 12/12 |

Rapport du worktree :
`artifacts/coverage/debug-complete-validated/322f81ab-0bb1-4fbd-b8eb-8959ac228321/coverage.cobertura.xml`.

## Procedures et options natives (lot après ab67e7b)

Matrices ajoutées avant première exécution : Procedures couvre les scalaires .NET, culture, identité/projet/module/déclaration/SHA, body hors bornes, plafond 2048, contexte absent/Post rejeté, requête capturée, queued/delivering/delivered/failed, rétention20 et transport natif validant commande absente avant lookup hôte. Options couvre champs requis, catalogue vide/9tabs/nom vide/2001controls, version, tab absent/ambigu, contrôle absent/ambigu/illisible, chaque type et allowlist localisée, readback absent/ambigu/error/type/visible/enabled et annulation après OK échoué ou dialogue resté ouvert. NativeOptionsProbe utilise vrais fournisseurs UIA jetables pour checkbox/radio/edit, patterns absents, indeterminate/read-only/password, candidats filtrés et ambigus, bouton OK absent/classe/disabled/Post refusé et click réussi. Seule frontière production ajoutée : IsWindowEnabled du bouton Options, défaut PInvoke inchangé ; tous delegates/contextes restaurés en finally/Dispose.

Validation finale : **92 tests verts, zéro ignoré**, build sans warning ni erreur.
- VbeDebug.Procedures.cs : 100/100 lignes, 162/162 branches.
- VbeDebugWindows.Options.cs : 78/78 lignes, 118/118 branches.
- NativeOptionsProbe dans VbeDebugWindows.cs : 112/112 lignes, 74/74 branches.
Rapport worktree : artifacts/coverage/procedures-options-validated/ce5b8900-658f-4bd6-9ce1-3041af3822f7/coverage.cobertura.xml.
Le contrôle de nom projet unique est exercé via sélecteur FileName exact : deux projets peuvent partager un Name tout en ayant des paths différents ; aucune garde retirée. Le contexte déterministe est retiré avant l'await du test de validation du transport pour éviter une continuation pendante de fixture. Exclusion collector unique [ProviderTests]*, aucune exclusion de production, aucune instance Excel/SOLIDWORKS contactée.
