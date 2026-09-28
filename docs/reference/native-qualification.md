# Qualification des cinq points VBE

État du 28 septembre 2026, branche `codex/chat-ux`. Ce lot traite les cinq validations natives annoncées après les extensions fonctionnelles. Il ne transforme pas les huit fonctions absentes ou bornées de [functional-extensions.md](functional-extensions.md) en fonctions complètes. Le thème natif VBE reste exclu.

## Résultats exacts

| Point | Preuve obtenue | Ce qui reste |
| --- | --- | --- |
| 1. Appels paramétrés depuis le complément installé | Excel isolé, PID et MVID du complément vérifiés. Sub avec chaîne française contenant des guillemets, Double, Boolean et Null ; Function appelée avec 21, résultat 42 relu indépendamment dans la feuille. | Objets, arguments nommés et retour structuré restent hors contrat ; variantes d’hôtes non qualifiées. |
| 2. Préférences et barres persistantes | Six cases Éditeur, tabulation aux bornes 1 et 32, trois choix d’interruption sur erreurs, deux cases de compilation : **13 essais** avec réouverture, relecture et restauration du hash de toutes les options. Barre et bouton natif ID 186 restaurés après fermeture/reprise d’Excel ; identité exacte du contrôle conservée par Copy. Position flottante (320, 220) également restaurée et relue après redémarrage. | Préférences natives anglaises et autres hôtes non exécutés. La restauration par profil concerne uniquement les barres VBAi explicitement persistantes ; elle ne personnalise pas les menus/boîtes à outils. |
| 3. SWP, rechargement et signatures | Aucun nouvel essai SOLIDWORKS effectué dans ce lot. | **En attente du fichier jetable `.swp` autorisé**. Save/SaveAs, rechargement, SHA et signature après persistance restent non qualifiés. Le contrôle hors ligne d’un certificat ne valide pas le digest de la signature VBA. |
| 4. UserForms, propriétés, événements et ActiveX | 14 types MSForms : **34 mutations/restaurations** (Left, Enabled et Caption quand exposée), relecture native indépendante ; huit Caption absentes/non lisibles pour leur type. **60 dispositions** sauvegardées et réouvertes sur UserForm, Frame et page MultiPage. Cycle modal/modeless avec QueryClose annulé puis accepté, état du parent et source vérifiés. | Les neuf MSComctl présents au catalogue sont tous **HOSTING_REFUSED**, zéro contrôle restant, avec « sujet non approuvé ». Aucune politique de confiance modifiée. Les autres propriétés et événements ne sont pas tous qualifiés ; catalogue de propriétés ≠ setter natif validé. |
| 5. Grands arbres, langues, DPI et ancrage | Tableau Excel de **1 000 Long**, expansion/repli, sept lignes scalaires (Double, Currency, Date, String, Boolean, Variant/Null, Long), reprise et marqueur 3000 ; source inchangée. **52 captures** de quatre fenêtres dans 13 cultures ; **8 captures** des quatre fenêtres sur deux écrans réels, à 96 DPI. | Arbres COM, espions volumineux, débogueur SOLIDWORKS et autres langues natives VBE non qualifiés. Les captures WinForms ne valident pas la qualité linguistique complète, le DPI 125/150/200 % ou l’ancrage natif VBE. Certaines chaînes restent en anglais par repli du catalogue. |

Ce tableau décrit les scénarios exécutés. Aucun de ces résultats ne prouve une couverture universelle de toutes les combinaisons VBA/COM.

## Corrections révélées par ces scénarios

- `Temporary=false` ne conservait pas effectivement la barre au redémarrage d’Excel : profils SQLite privés dans `%LocalAppData%/VBAi/VbeToolbars/<HOST>.sqlite`, fusion transactionnelle, bornes et suppression par identité.
- Ajouter un bouton par ID Office créait « Macros » au lieu du contrôle VBE « Exécuter » : copie du contrôle natif existant, ID/type/possession vérifiés, libellé de restauration dépendant du contexte.
- Les libellés français réels de complétion, info express, compilation sur demande et largeur de tabulation différaient de la liste autorisée. Le champ de tabulation partage aussi son nom avec une étiquette ; seul le type éditable est sélectionné.
- Une séquence de mutations MSForms suivie de l’ajout d’un contrôle a provoqué une corruption de heap Excel (`0xc0000374`). Le membre fautif n’a pas été isolé individuellement. Les propriétés scalaires usuelles utilisent désormais un dispatch COM typé ; la matrice de 34 mutations passe avec cette correction. Les autres gardes restent actives.
- UIA exposait deux fois chaque enfant du tableau. La capture retire uniquement les observations identiques par texte complet et chemin ; **1 000 doublons retirés, 1 000 valeurs conservées**, dernière valeur 3000. Des valeurs/contextes/chemins distincts sont conservés. Le contrat reste `UIAExposedRowsOnly`.

### Gardes MSForms conservées

Écriture refusée pour `ComboBox.ColumnCount`, `TextBox.ScrollBars`, `SpinButton.Min/Max/Value/Delay/SmallChange`, `ToggleButton.Value`, et `Label.Cancel` sans setter utilisable. `_Font_Reserved` est un membre COM réservé indisponible. Ces restrictions ne sont pas levées par les essais des propriétés usuelles.

## Reproduire et consulter les preuves

Les scénarios refusent les hôtes VBA déjà ouverts, identifient leur propre processus et utilisent un classeur jetable. L’enregistrement COM temporaire et AccessVBOM sont restaurés ; aucune macro utilisateur n’est exécutée. Le wrapper exige les opt-in explicites `AllowTemporaryRegistration` et `AllowTemporaryVbaAccess`.

- `tools/tests/Test-RegisteredExcelNavigation.ps1`, scénario `FunctionalExtensions` : appelle `tools/probes/Test-RegisteredFunctionalExtensions.ps1` ; preuves finales dans `artifacts/native-qualification/registered-final/` (rapport, matrice des options, propriétés, ActiveX et nettoyage).
- Même wrapper, scénario `DebuggerTree` : appelle `tools/probes/Test-RegisteredDebuggerTree.ps1` ; `artifacts/native-qualification/debugger-tree-final/registered-debugger-tree.json`.
- Scénario `FormLifecycle` : `artifacts/native-qualification/form-lifecycle/installed-form-lifecycle.json`.
- `tools/tests/Test-ExcelFormLayouts.ps1` : preuve dans `artifacts/native-qualification/form-layouts/`.
- `tools/tests/Test-MultilingualWindows.ps1` : captures dans `artifacts/localization/multilingual-windows/`.
- `tools/tests/Test-WinFormsDisplayProfiles.ps1` : mesures/captures dans `artifacts/native-qualification/displays/` ; DPI du processus et écran effectivement utilisés enregistrés.

Les artefacts sont locaux et ignorés par Git ; les scripts et ce bilan sont versionnés. Aucun fichier Designer ou agencement WinForms n’a changé dans ce lot. Aucun pourcentage de couverture n’a été remesuré.

## Conditions nécessaires pour terminer les qualifications restantes

1. Fichier SOLIDWORKS `.swp` jetable identifié et autorisé pour modification/sauvegarde, puis essai réel de rechargement et de signature.
2. Contrôles ActiveX tiers autorisés par l’hôte pour tester leurs setters/événements ; les refus actuels sont des résultats de qualification, pas des succès d’hébergement.
3. Profils natifs VBE de langues supplémentaires et écrans/DPI différents, puis ancrage du chat vérifié dans ces profils.
4. Matrice de propriétés/événements supplémentaire et arbres COM/espions représentatifs. Les gardes qui empêchent les mutations précédemment corrélées à des crashs restent des fonctions limitées.

## Compilation et suite locale finale

Compilation : **0 erreur, 0 avertissement**. Suite Unit complète : **909 réussis, 0 échec, 0 ignoré**, rapport `artifacts/native-qualification/tests/unit-final.trx`. Ce passage Unit ne remplace pas les essais hôtes absents. Les 1 000 valeurs du rapport du débogueur ont aussi été contrôlées individuellement contre la fixture `index * 3`. Enregistrement COM, AccessVBOM et nettoyage du scénario final vérifiés restaurés (`registered-final/cleanup.json`), aucune instance Excel/SOLIDWORKS laissée ouverte.
