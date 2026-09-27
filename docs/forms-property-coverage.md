# Propriétés du concepteur MSForms : état de couverture

L'inventaire source [excel-control-properties.csv](excel-control-properties.csv) contient 771 descripteurs réels relevés sur les 14 contrôles natifs dans Excel. `form_tree` lit les propriétés des contrôles, Frames, Pages et Tabs par chemin canonique ; `form_control_properties` lit les contrôles du premier niveau. Les descripteurs n'établissent pas à eux seuls qu'un setter COM fonctionne.

| Type | Descripteurs | Getters en erreur | Lecture seule déclarée | Descripteur modifiable hors erreur |
| --- | ---: | ---: | ---: | ---: |
| CheckBox | 57 | 1 | 9 | 47 |
| ComboBox | 86 | 1 | 16 | 69 |
| CommandButton | 49 | 1 | 7 | 41 |
| Frame | 63 | 1 | 15 | 47 |
| Image | 38 | 0 | 7 | 31 |
| Label | 51 | 1 | 7 | 43 |
| ListBox | 64 | 1 | 10 | 53 |
| MultiPage | 45 | 1 | 9 | 35 |
| OptionButton | 57 | 1 | 9 | 47 |
| ScrollBar | 38 | 0 | 7 | 31 |
| SpinButton | 36 | 0 | 7 | 29 |
| TabStrip | 51 | 1 | 13 | 37 |
| TextBox | 79 | 1 | 14 | 64 |
| ToggleButton | 57 | 1 | 9 | 47 |
| **Total** | **771** | **11** | **139** | **621** |

Les 11 getters en erreur sont tous `_Font_Reserved` (`DISP_E_MEMBERNOTFOUND`). Ce nom reste visible dans le descripteur COM, mais sa valeur n'est pas lisible dans l'Excel testé. Les 621 descripteurs marqués modifiables sont un **plafond de possibilités à qualifier**, pas 621 écritures validées : `Label.Cancel` annonçait un setter qui a échoué, et certaines écritures sur ToggleButton/SpinButton ont précédé des crashs Excel. Les garde-fous correspondants sont décrits dans [forms-duplication-coverage.md](forms-duplication-coverage.md).

La lecture vivante expose maintenant un `SetterStatus` distinct pour les cas connus : `GetterUnavailable` pour `_Font_Reserved`, `BlockedNativeSetterFailure` pour `Label.Cancel`, `BlockedAfterHostCrash` pour `ToggleButton.Value`, `TextBox.ScrollBars`, `ComboBox.ColumnCount` et `SpinButton.Min/Max/Value/Delay/SmallChange`. Les 612 autres descripteurs annoncés modifiables restent `DescriptorCandidateUnverified` jusqu'à preuve de mutation et relecture sur le type exact. Les 139 `DescriptorReadOnly` décrivent seulement la métadonnée COM. Le code refuse aussi `_Font_Reserved` avant tout setter ; les refus des propriétés à risque ont été testés séparément pour ToggleButton/SpinButton. `Test-FormPropertyStatus.ps1` dans Excel PID 28636 a comparé huit statuts entre `form_tree` et `form_control_properties`, vérifié le refus `_Font_Reserved` sans changement de `TreeVersion`, puis Excel s'est fermé sans événement Application Error 1000.

Qualifications ciblées du 27 septembre 2026 : `Label.Font.Italic` a été relu `False→True` dans Excel PID 13440, `Label.ForeColor` a été relu `Color [ControlText]→Color [Red]` dans PID 49172 ; les versions ont changé puis sont restées stables, et les deux sessions se sont fermées sans Event1000. En revanche, une seule écriture `TextBox.ScrollBars=Vertical` a été relue `None→Vertical` dans PID 41684, puis Excel s'est arrêté avant la fermeture prévue. Application Error 1000 à 04:20:50 : `ntdll.dll`, `0xc0000374`. Cette propriété est bloquée avant COM jusqu'à isolation de la cause ; aucune autre écriture de la famille enum n'est déduite de cet essai.

Parmi ces descripteurs, 110 propriétés sont des énumérations. Leur nom de type COM contient un préfixe numérique propre à la session ; l'identifiant ne doit pas être persisté. `AllowedValues` expose maintenant les noms de choix depuis le type vivant dans `form_tree`, `form_control_properties` et `form_properties`. Dans Excel PID 40884, `Test-ControlEnumChoices.ps1` a relu les 14 types, 18 nœuds, les 771 propriétés et 110 catalogues enum sans choix manquant. `Label.TextAlign` a rendu `Left`, `Center`, `Right`. Deux lectures consécutives ont gardé la même `TreeVersion` ; Excel s'est fermé normalement sans événement Application Error 1000.

Les prochaines qualifications doivent séparer les **valeurs lisibles**, les **setters effectivement relus**, les **objets exigeant une commande dédiée** (`Font`, `Picture`, `MouseIcon`, collections) et les propriétés vraiment non exposées ou lecture seule. Les 30 `Com2Variant`, 42 `Font`/image/icône et 33 couleurs demandent une conversion ou un flux propre ; certains cas sont déjà prouvés ponctuellement, sans validation de toute leur famille. Les Pages et Tabs, ainsi que les contrôles ActiveX additionnels, exigent leurs propres inventaires.

`ComboBox.List` et `ListBox.List` sont des propriétés **indexées** absentes des 771 descripteurs scalaires. `form_list_items` lit au plus 128 cellules par appel, avec pagination par ligne, chemin canonique et erreur propre à chaque cellule. Dans Excel jetable PID 36020, elle a lu les deux contrôles vides (`ListCount=0`, `ColumnCount=1`) et refusé un chemin absent. Dans PID 16640, une sonde bridge-only a ajouté cinq chaînes avec `AddItem` à chacun des deux contrôles ; la lecture a restitué `Alpha`, `Beta`, `Gamma`, `Delta`, `Epsilon` en pages de 2, 2 et 1 ligne, avec `HasMore` correct et `TreeVersion` stable entre lectures. Excel s'est fermé normalement, sans Event1000, et le CodeBase COM a été restauré. La lecture est exposée au LLM ; la mutation de sonde ne l'est pas. Le cas multicolonne et les listes liées à `RowSource` restent non vérifiés. `TreeVersion` peut changer avec `ListCount` mais n'est pas une empreinte des valeurs indexées.

Essai multicolonne du 27 septembre 2026 : Excel jetable PID 49796 a reçu une ComboBox, `ColumnCount=2`, puis trois `AddItem`. `form_list_items` a correctement relu les trois chaînes en première colonne et les trois cellules vides de la seconde, avec pagination 2+1. Excel a toutefois disparu **avant la fermeture** ; Event1000 à 04:47:02, `ntdll.dll`, `0xc0000374`. La cellule indexée de seconde colonne et `RowSource` n'ont **pas** été écrits. La cause exacte dans la séquence `ColumnCount` puis `AddItem` n'est pas isolée : `ComboBox.ColumnCount` est bloqué avant COM, et la sonde `AddItem` refuse désormais les listes multicolonnes. Aucun essai `RowSource` n'a été effectué. La lecture de listes déjà multicolonnes a produit un résultat, mais ce cycle ne constitue pas une preuve de stabilité. Le garde a été vérifié dans Excel jetable PID 35568 : `SetterStatus=BlockedAfterHostCrash`, refus avant COM, version inchangée, fermeture normale, zéro Event1000.

Lecture seule de WER : les crashs PID 49796 et 41684 ont tous deux un dump local d'environ 66 Mo sous `%LOCALAPPDATA%\CrashDumps\EXCEL.EXE.<PID>.dmp` et un `Report.wer` dans `C:\ProgramData\Microsoft\Windows\WER\ReportArchive`. Les deux rapports ont le même `Response.BucketId=5f1dc203faf580d9d2edcca5973ec490`, l'étiquette `OFFICE_MODULE_VERSION_MISMATCH`, et chargent `FM20.DLL`/`fm20FRA.DLL`. Event1000 donne dans les deux cas `ntdll.dll` et `0xc0000374`. L'étiquette WER et la présence de FM20 ne prouvent pas que le complément ou FM20 est fautif. Aucun WinDbg/CDB ou autre analyseur de dump n'a été trouvé installé ; les piles natives n'ont pas été décodées. La cause native reste indéterminée et les mutations concernées ne doivent pas être réessayées sans diagnostic adapté.
