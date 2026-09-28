# Recherche : personnalisation native du VBE

Date : 28 septembre 2026. Recherche demandée par l’utilisateur, confiée à un agent ; sources primaires consultées. Aucun code tiers copié.

Suite : [étude de stabilisation du rendu](native-theme-rendering-strategy.md), réalisée après les signalements de retour au clair et de texte fin. Elle distingue les contrats de dessin Microsoft, les limites du post-traitement des pixels et le prototype proposé avant une évolution du moteur.

## Conclusion

Rubberduck fournit des mécanismes de suivi et de sous-classement des fenêtres VBE, mais aucun moteur de thème sombre natif des menus/dialogues n’a été identifié dans les fichiers examinés. Ses panneaux modernes et ses commandes ajoutées ne prouvent pas qu’il remplace le dessin des contrôles Office.

## Sources et portée

| Source | Constat vérifié | Application au projet |
| --- | --- | --- |
| [Rubberduck SubclassManager](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/WindowsApi/SubclassManager.cs), [CodePaneSubclass](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/WindowsApi/CodePaneSubclass.cs) | Sous-classes du code et du designer ; saisie/focus, pas de moteur de peinture sombre | Référence de cycle de vie uniquement |
| [VBENativeServices](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/Events/VBENativeServices.cs) | Hook OutOfContext filtré sur le thread VBE ; code identifié par VbaWindow + ObtbarWndClass | Vérifier le thread de nos callbacks et fenêtres |
| [DockableWindowHost](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Main/ComClientLibrary/UI/DockableWindowHost.cs) | Conteneurs VBFloatingPalette / DockingView ; hébergement ActiveX des panneaux propres | Examiner Propriétés en mode flottant, avec parent et propriétaire réels |
| [AppCommandBarBase](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Core/UI/Command/MenuItems/CommandBars/AppCommandBarBase.cs) | Ajout de boutons COM et mise à jour par dispatcher | Ne constitue pas une solution de dessin des menus Office |
| [Notepad++ NppDarkMode.cpp](https://github.com/notepad-plus-plus/notepad-plus-plus/blob/master/PowerEditor/src/NppDarkMode.cpp) | Traitements distincts Button, Groupbox, Tab, ComboBox et cleanup WM_NCDESTROY | Privilégier des corrections adaptées à chaque contrôle |
| [VBEThemeColorEditor](https://github.com/gallaux/VBEThemeColorEditor) | Patch de la palette de 16 couleurs dans la DLL VBE | Ne traite pas nos menus/dialogues ; aucun patch de DLL retenu |
| [VBE-Themes](https://github.com/vicsar/VBE-Themes) | Configuration des couleurs de code dans le registre | Ne résout pas le dessin des fenêtres natives |

Les sources GitHub utilisent des branches évolutives ; ces constats décrivent la consultation à la date indiquée. Rubberduck est sous GPLv3 : réutiliser des principes documentés n’autorise pas à copier son code sans examiner les obligations de licence. Ses signatures interop ne doivent pas être reprises aveuglément : conserver LRESULT/IntPtr en x64.

## Points techniques à vérifier

1. [SetWindowSubclass](https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass) ne permet pas le sous-classement entre threads. [SetWinEventHook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwineventhook) OutOfContext livre les callbacks sur le thread installateur. Mesurer ces identités avant de changer le dispatch.
2. Les menus Notepad++ utilisent WM_UAHDRAWMENU/WM_UAHDRAWMENUITEM et HMENU. Aucune preuve d’applicabilité aux MsoCommandBar Office. Relever HWND, classe, parent/propriétaire, PID/TID et messages d’un vrai menu VBE ouvert.
3. La fenêtre Propriétés flottante peut changer de conteneur. La présence de classes dans Rubberduck constitue une piste, pas une preuve de notre défaut.
4. [win32-darkmode](https://github.com/ysc3839/win32-darkmode/blob/master/win32-darkmode/DarkMode.h) distingue les signatures de l’ordinal135 selon la version Windows. Notre support doit être borné ; pas de patch IAT transposé dans SOLIDWORKS.
5. Le support Office/VB6 annoncé par [Rubberduck](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing) ne prouve pas la compatibilité avec SOLIDWORKS 2020.

## Contrôle local lié à cette recherche

La sonde identifie désormais le dialogue par sa création, son PID, sa chaîne de propriétaire et sa classe #32770. Les artefacts `vbe-native-dialog-identified`, `vbe-native-dialog-contrast` puis `vbe-native-dialog-controls` montrent successivement le défaut de texte noir, sa correction, puis les boutons de commande sombres conservés. Les libellés des cases/groupes utilisent le rendu classique et WM_CTLCOLOR ; les boutons de commande conservent leur thème natif. Les onglets des dialogues sont traités séparément. Les trois essais ont terminé avec code 0 et Excel fermé ; aucun changement de palette de code n’était demandé.

La dernière capture est celle d’Options seulement : la recherche de commande Propriétés n’a pas identifié de candidat. La barre de titre reste claire dans PrintWindow ; une preuve à l’écran reste nécessaire. Les menus supérieurs et Propriétés flottante restent ouverts au suivi, ainsi que les interactions, la désactivation et SOLIDWORKS. Aucune finalisation globale n’est revendiquée.

## Suite expérimentale : menus Office identifiés

Après cette recherche, la sonde a identifié `MsoCommandBarPopup` comme classe des menus déroulants, distincte de `MsoCommandBar`. Ajouter cette classe au traitement existant a produit des captures sombres reproductibles pour Fichier, Édition et Affichage sur deux ouvertures. Aucun traitement HMENU/WM_UAHDRAWMENU de Notepad++ n’a été transposé. Les limites et preuves figurent dans `native-dark-theme.md`.
