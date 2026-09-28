# Stabilisation du rendu natif du VBE

Date : 28 septembre 2026. Statut : orientation validée par l'utilisateur ; premier prototype en cours.

## Décision retenue pour le prototype

Remplacer progressivement la recoloration de l'image affichée par un traitement au moment du dessin de chaque famille de contrôles. Pour le code, faire dessiner les caractères avec leur couleur finale afin de conserver l'anticrénelage natif.

Le thème complet reste l'objectif : code et fenêtre Exécution, panneaux, Propriétés et ses onglets, menus, dialogues, persistance des couleurs et réversibilité. Les couleurs inspirées de Visual Studio Community VB.NET et le fond anthracite restent demandés. Une palette native limitée ou quelques panneaux stabilisés ne suffisent pas à déclarer ce travail terminé.

La recherche ne démontre pas encore l'existence d'un moteur complet prêt à intégrer dans le VBE de SOLIDWORKS. Elle établit les limites de notre approche et les expériences nécessaires pour choisir un remplacement.

## 1. Ce que notre code fait actuellement

Audit de `src/CodexVBE/Ui/VbeNativeChrome.cs` et `VbeNativeTheme.cs` dans le répertoire courant du projet :

| Constat dans le code | Conséquence ou risque |
| --- | --- |
| `Paint` prend un GetDC/GetWindowDC, copie les pixels avec BitBlt, les transforme puis recopie l'image entière. | La source peut mélanger un dessin natif frais et des pixels déjà convertis. Sa provenance n'est pas suivie. |
| La conversion intervient après DefSubclassProc, puis lors d'un message différé. Certains messages programment à nouveau tous les contrôles suivis. | Cette succession ne garantit pas que notre passe soit le dernier dessin du VBE. Elle augmente aussi le travail par rafraîchissement. |
| `HasDarkBackground` décide du traitement de toute une surface à partir de quatre pixels. | Texte, sélection ou indicateur sur ces positions peuvent modifier la décision pour toute la fenêtre. |
| `MapPixel` traite les couleurs restantes comme des niveaux de luminance ; `MapCodePixel` moyenne certaines couvertures RGB. | L'information de sous-pixel ClearType est perdue. Des couleurs sémantiques, telles que les échantillons dans Propriétés, peuvent aussi être transformées. |
| WM_PRINT emprunte des chemins client et non-client avec le même HDC sans traitement complet des drapeaux PRF_* et des origines. | Une capture PrintWindow peut différer de l'affichage réel et ne constitue pas une qualification suffisante. |
| Les hooks de fenêtres sont filtrés par processus, sans vérification explicite du thread propriétaire lors du sous-classement. | Le respect des contraintes de thread doit être vérifié, notamment pour les fenêtres auxiliaires. |

Ce sont des observations de code. Elles ne prouvent pas individuellement la cause de chaque anomalie observée.

Preuves locales complémentaires : `artifacts/solidworks-native-repaint` montre un titre Projet clair avant rafraîchissement natif et sombre après. La capture est partiellement masquée ; la conclusion porte sur le titre visible. Les pixels de bord du code contiennent des teintes sombres restées sur la palette d'origine. Cela est cohérent avec le texte fin signalé, sans constituer une mesure complète de la police ou de tous les DPI.

## 2. Enseignements Microsoft

### Peinture et régions invalides

WM_PAINT et BeginPaint organisent une transaction liée à la région invalide. Un GetDC utilisé après cette transaction ne reprend pas automatiquement sa région de peinture. Le traitement devrait respecter les zones concernées, le contexte graphique fourni et l'état de dessin. Sources : [WM_PAINT](https://learn.microsoft.com/en-us/windows/win32/gdi/wm-paint), [BeginPaint](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-beginpaint).

### Couleurs avant rasterisation

SetTextColor définit la couleur employée par TextOut/ExtTextOut. Les contrôles qui proposent NM_CUSTOMDRAW permettent de fournir leurs couleurs avant le dessin. C'est le point d'intervention à privilégier pour conserver leur comportement et leur rendu de texte. Sources : [SetTextColor](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-settextcolor), [Using Custom Draw](https://learn.microsoft.com/en-us/windows/win32/controls/using-custom-draw).

ClearType tient compte du fond sur lequel il dessine. Une transformation de l'image après calcul des contours n'est donc pas équivalente au dessin initial sur le fond sombre final. Source : [Microsoft — ClearType et fond fixe](https://devblogs.microsoft.com/oldnewthing/20150129-00/?p=44803).

### Limites du double tampon

Un tampon peut rendre une présentation atomique et limiter le scintillement. Il ne rétablit pas l'information ClearType supprimée par une conversion et ne garantit pas que tous les dessins du VBE passent par lui. WM_PRINTCLIENT demande au contrôle de dessiner dans un HDC fourni ; son exploitation suppose que ce contrôle implémente correctement le message. Le seul succès de PrintWindow ne prouve pas cette couverture. Sources : [WM_PRINTCLIENT](https://learn.microsoft.com/en-us/windows/win32/gdi/wm-printclient), [BeginBufferedPaint](https://learn.microsoft.com/en-us/windows/win32/api/uxtheme/nf-uxtheme-beginbufferedpaint).

### Portée des thèmes Windows

Le mode sombre DWM documenté concerne notamment le cadre standard de la fenêtre. Il ne suffit pas à refaire les surfaces Office personnalisées. Les options VBA documentées donnent accès aux couleurs des catégories du code et à la police ; elles ne constituent pas un moteur de thème complet. Sources : [thèmes Win32](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/ui/apply-windows-themes), [options de l'environnement VBA](https://learn.microsoft.com/fr-fr/office/vba/language/how-to/set-visual-basic-environment-options).

SetWindowSubclass impose aussi de rester sur le thread de la fenêtre. La gestion des handles, destructions et réutilisations doit être explicite. Source : [SetWindowSubclass](https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass).

## 3. Projets GitHub examinés

| Projet | Ce que fait réellement le code | Transfert possible et limite |
| --- | --- | --- |
| [win32-darkmodelib — DmlibSubclassControl.cpp](https://github.com/ozone10/win32-darkmodelib/blob/25298e8a92ed0fac8e17b3e6878df5bbd7a5d699/src/DmlibSubclassControl.cpp) | Dessine boutons, onglets et ComboBox avec leurs couleurs et états finaux. Certains chemins utilisent un tampon. Traite aussi des dessins directs lors de WM_ENABLE et libère les ressources au changement de thème/DPI. | Référence la plus pertinente pour nos contrôles standards. Les onglets sont dessinés à partir de leurs données TCITEM et de leur état. Ce fichier est MPL-2.0 ; le dépôt contient aussi des fichiers MIT. Aucun support VBE/SOLIDWORKS démontré. |
| [win32-darkmodelib — DmlibHook.cpp](https://github.com/ozone10/win32-darkmodelib/blob/25298e8a92ed0fac8e17b3e6878df5bbd7a5d699/src/DmlibHook.cpp) | Certains contrôles utilisent des interceptions IAT dans comctl32/comdlg32/uxtheme. | Leur portée peut dépasser les fenêtres VBE. Une intégration globale dans SOLIDWORKS serait inappropriée sans isolation démontrée. |
| [Rubberduck — services natifs](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/Events/VBENativeServices.cs) | Identifie et suit les fenêtres VBE ; organise les événements et le sous-classement. | Utile pour leur durée de vie et l'affinité de thread. Aucun moteur de thème natif complet identifié dans les fichiers examinés. |
| [Notepad++ — NppDarkMode.cpp](https://github.com/notepad-plus-plus/notepad-plus-plus/blob/master/PowerEditor/src/NppDarkMode.cpp) | Traitements séparés par contrôle et couleurs choisies avant dessin. | Référence de conception ; code GPLv3+. Ses menus HMENU/UAH ne sont pas nos popups MsoCommandBar. |
| [VBEThemeColorTool](https://github.com/furyutei/VBEThemeColorTool), [implémentation](https://github.com/furyutei/VBEThemeColorTool/blob/master/src/vbetctool.py) | Modifie deux tableaux de 16 couleurs dans VBE7.DLL et les affectations du registre. | Agit réellement avant le rendu du code, mais modifie une DLL installée et dépend de sa structure. Le README signale des problèmes de l'ancien patch, dont des crashes Options dans Excel x64, et les effets des mises à jour Office. Pas de qualification SOLIDWORKS. |
| [VBE-Themes](https://github.com/vicsar/VBE-Themes) | Configure les indices de couleur dans le registre. | Ne suffit pas à obtenir les RGB souhaités sans modifier la palette sous-jacente ; ne couvre pas le chrome. |
| [xlide_vbide](https://github.com/WilliamSmithEdward/xlide_vbide), [HostChrome.cs](https://github.com/WilliamSmithEdward/xlide_vbide/blob/f97d02bc5b37dac98e265e3c39cf2cf66fc83f29/src/Xlide.Vbe.Shim/Editor/HostChrome.cs) | Place un éditeur Monaco/WebView à la place de la surface document, tout en conservant le moteur VBE. | Piste pour le futur éditeur moderne, hors de la finalisation de l'habillage natif actuelle. |

Une adaptation indépendante des mécanismes documentés évite de recopier du code sans analyse de licence. Les références de branches restent évolutives ; le fichier darkmodelib cité est fixé à son commit consulté.

Une modification de palette uniquement en mémoire serait également une piste de recherche, mais aucune implémentation publique VBE qualifiée n'a été identifiée dans cette étude. Elle resterait dépendante des internes du runtime. Le patch de DLL sur disque n'est pas proposé pour notre architecture.

## 4. Architecture recommandée

| Surface | Voie privilégiée | Point à démontrer |
| --- | --- | --- |
| Arbres, listes et champs standards | Messages de couleur et custom draw documentés, selon le contrôle. | États sélectionné, désactivé, focus, création et destruction. |
| Bordures, titres et onglets | Dessin limité aux parties dont nous maîtrisons les rectangles et les états, au bon moment de leur peinture. | Repaint partiel, activation, masquage, flottement et DPI. |
| Grille Propriétés | Intervention dans le dessin de la ligne, avec couleurs appliquées avant le texte lorsque le moteur le permet. | Le VBE est propriétaire du rendu ; WM_DRAWITEM ne fournit pas nécessairement des valeurs textuelles exploitables. Préserver les pastilles de couleur. |
| Code et Exécution | Palette native pour référence, puis prototype de couleurs appliquées avant la rasterisation pour atteindre la palette souhaitée. | Le chemin de dessin exact, les HDC mémoire, les sélections, le curseur et les marqueurs du débogueur. |
| MsoCommandBar et MsoCommandBarPopup | Backend spécifique après observation de leurs appels de dessin. | Aucun transfert automatique des recettes HMENU/WM_UAHDRAWMENU. |

Le gestionnaire commun devrait suivre chaque HWND, son thread, sa classe, sa durée de vie et son backend. Les rafraîchissements seraient limités au contrôle concerné. Les gestionnaires de peinture éviteraient les appels COM, le pompage de messages, l'ouverture de dialogues et les allocations d'images plein écran à chaque passage.

### Surfaces sans extension de dessin exploitable

Une interception native ciblée des appels graphiques est une piste pour appliquer texte et fond avant leur rendu. Microsoft Detours fournit un mécanisme d'interception en mémoire et prend en charge x64. Cela ne démontre pas son applicabilité au VBE : tous les appels à une fonction détournée passent par le detour, et Microsoft ne garantit pas le logiciel ainsi modifié. Sources : [Detours](https://github.com/microsoft/Detours), [fonctionnement](https://github.com/microsoft/Detours/wiki/OverviewInterception), [contraintes d'utilisation](https://github.com/microsoft/Detours/wiki/Using-Detours).

Un prototype doit donc identifier précisément le contexte VBE et les appels utiles. Un hook générique de GetSysColor ou SetTextColor à l'échelle de SOLIDWORKS serait trop large. Les HDC mémoire, les appels hors WM_PAINT et les caches Office rendent une simple condition sur le HWND ou le thread insuffisante. Le premier prototype devrait tracer les appels et leurs contextes sans modifier les couleurs. Aucune interception n'a été installée dans le cadre de cette recherche.

## 5. Jalon autorisé

1. Réaliser un premier contrôle pilote sur les onglets de Propriétés : leurs textes et états sont accessibles par les API standard. Dessiner avec les couleurs finales et retirer la conversion d'image pour ce seul contrôle.
2. En parallèle, construire une sonde de traçage bornée sur une fenêtre de code et une ligne de Propriétés, dans une instance de test. Identifier le chemin effectif avant toute modification des couleurs ; comparer avec le texte natif à police et DPI identiques.
3. Choisir leur backend d'après cette trace : couleur/custom draw direct s'il existe ; interception ciblée si le périmètre est prouvé ; tampon seulement si le dessin complet hors écran est réellement supporté. Prototyper avec l'ancienne conversion désactivée pour les surfaces prises en charge. La police et sa taille restent identiques.
4. Vérifier le comportement réel : saisie et sélection, défilement, masque/réapparition, menu ouvert/fermé, fenêtre déplacée/redimensionnée, fond/inactif, onglets de Propriétés, breakpoint et pas à pas, DPI disponibles, désactivation et fermeture normale de l'hôte.
5. Étendre seulement après un résultat probant. Les mesures doivent inclure l'absence de recoloration dans SOLIDWORKS hors VBE, l'absence de fuite de handles GDI et l'absence d'activité de repaint persistante au repos.

La réussite attendue est un dessin net et stable sur ces parcours, avec des couleurs sémantiques préservées. Un test unitaire de conversion RGB ou une capture identique ne suffiront pas à qualifier le moteur.

## 6. État du projet à la fin de la recherche

Aucun nouveau code de rendu ni aucune DLL hôte n'ont été modifiés pour cette recherche. Les derniers correctifs du dialogue Options et des titres sont dans la DLL installée. Le correctif expérimental des pixels ClearType est compilé séparément et n'a pas été installé avant l'interruption : les empreintes de `bin/Debug/net48/CodexVBE.dll` et `artifacts/native-palette-dialog-check/CodexVBE.dll` diffèrent. La recherche recommande de réexaminer ce type de correction dans le cadre du changement de pipeline décrit ci-dessus.

## 7. Mise en œuvre du premier prototype

Après accord de l'utilisateur, le pilote des onglets Propriétés utilise une classe dédiée, `VbeNativePropertyTabs`. Elle lit les deux libellés et leurs rectangles natifs, conserve la police du contrôle et applique les couleurs GDI avant `DrawTextW`. La peinture écran passe par `BeginPaint`/`EndPaint` ; l'impression utilise le HDC fourni et préserve son état. Aucun bitmap de l'écran n'est relu pour ces onglets. Les autres surfaces conservent provisoirement leur moteur existant.

Le gestionnaire exclut ces HWND des anciennes passes de recoloration, y compris des messages différés déjà programmés. Il vérifie aussi le thread propriétaire avant le sous-classement et réessaie l'attachement après l'initialisation des deux onglets natifs. Les dispositions hors périmètre (verticales, multi-lignes, images, owner-draw ou RTL) ne sont pas revendiquées par ce pilote. Un changement ultérieur vers une disposition non prise en charge rend la peinture au contrôle natif.

La sonde C++ séparée, sous `tools/probes/native-render-trace`, vise les imports graphiques du module VBE7 chargé. Elle doit enregistrer les appels et leurs contextes sans le contenu du code, sans modifier les couleurs, avec une limite d'événements et une restauration des slots IAT à l'arrêt. Elle appartient aux outils de diagnostic et n'est pas une dépendance distribuée de l'add-in. Les imports observés sur les VBE7 locaux incluent `SetTextColor`, `SetBkColor`, `TextOutA`, `ExtTextOutA`, `CreateCompatibleDC` et `BitBlt`.

Les résultats de compilation, des contrôles natifs isolés et du passage Excel sont consignés dans le [bilan du pilote](native-theme-renderer-pilot.md). La trace réelle confirme des dessins du code hors WM_PAINT. Une capture ciblée confirme aussi le clignotement des barres ; la limitation expérimentale des rafraîchissements laisse des boutons clairs et reste désactivée par défaut. Ces résultats ne valent pas validation SOLIDWORKS, multi-DPI, ni finalisation du thème complet.
