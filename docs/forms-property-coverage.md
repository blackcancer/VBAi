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

La lecture vivante expose maintenant un `SetterStatus` distinct pour les cas connus : `GetterUnavailable` pour `_Font_Reserved`, `BlockedNativeSetterFailure` pour `Label.Cancel`, `BlockedAfterHostCrash` pour `ToggleButton.Value` et `SpinButton.Min/Max/Value/Delay/SmallChange`. Les 614 autres descripteurs annoncés modifiables restent `DescriptorCandidateUnverified` jusqu'à preuve de mutation et relecture sur le type exact. Les 139 `DescriptorReadOnly` décrivent seulement la métadonnée COM. Le code refuse aussi `_Font_Reserved` avant tout setter ; les refus des sept propriétés à risque ont été testés séparément pour ToggleButton/SpinButton. `Test-FormPropertyStatus.ps1` dans Excel PID 28636 a comparé huit statuts entre `form_tree` et `form_control_properties`, vérifié le refus `_Font_Reserved` sans changement de `TreeVersion`, puis Excel s'est fermé sans événement Application Error 1000.

Parmi ces descripteurs, 110 propriétés sont des énumérations. Leur nom de type COM contient un préfixe numérique propre à la session ; l'identifiant ne doit pas être persisté. `AllowedValues` expose maintenant les noms de choix depuis le type vivant dans `form_tree`, `form_control_properties` et `form_properties`. Dans Excel PID 40884, `Test-ControlEnumChoices.ps1` a relu les 14 types, 18 nœuds, les 771 propriétés et 110 catalogues enum sans choix manquant. `Label.TextAlign` a rendu `Left`, `Center`, `Right`. Deux lectures consécutives ont gardé la même `TreeVersion` ; Excel s'est fermé normalement sans événement Application Error 1000.

Les prochaines qualifications doivent séparer les **valeurs lisibles**, les **setters effectivement relus**, les **objets exigeant une commande dédiée** (`Font`, `Picture`, `MouseIcon`, collections) et les propriétés vraiment non exposées ou lecture seule. Les 30 `Com2Variant`, 42 `Font`/image/icône et 33 couleurs demandent une conversion ou un flux propre ; certains cas sont déjà prouvés ponctuellement, sans validation de toute leur famille. Les Pages et Tabs, ainsi que les contrôles ActiveX additionnels, exigent leurs propres inventaires.
