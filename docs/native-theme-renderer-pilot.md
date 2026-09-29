# Pilote du rendu natif — 28 septembre 2026

## Objectif et périmètre

Mettre en œuvre la [stratégie validée](native-theme-rendering-strategy.md) : dessiner les onglets Propriétés avec leurs couleurs finales et observer les appels graphiques du code avant de choisir son prochain moteur. Le thème complet, sa netteté et sa stabilité restent à finaliser.

## Onglets Propriétés

`VbeNativePropertyTabs` conserve la police et le fonctionnement du TabControl natif. Il lit ses libellés et rectangles, puis dessine le fond et le texte directement en GDI. Les anciennes conversions de bitmap sont exclues pour ces HWND. Les onglets de dialogues et les configurations non prises en charge conservent leur voie native ou antérieure selon leur attachement.

Le pilote est limité aux deux onglets textuels horizontaux observés dans le VBE. Images, owner-draw, multiligne, RTL, orientations alternatives et débordement avec contrôle enfant sont exclus. La configuration réelle observée est `0x54001040`. La police, sa taille et sa qualité ne sont pas remplacées.

### Vérifications effectuées

Quatre tests ciblés sur un véritable `SysTabControl32` ont réussi après correction de deux montages de test : fenêtre hors écran mais visible pour vérifier la région invalide ; combinaisons valides de styles d'alignement avec `TCS_FIXEDWIDTH`. Ils couvrent sélection et couleurs, police conservée, HDC fourni/clip/viewport, région de peinture, rejet d'autres threads/styles et ressources GDI. La suite globale n'a pas été exécutée.

Le passage réel `artifacts/native-renderer-pilot/excel-run-01` utilise un classeur Excel jetable :

| Contrôle | Résultat observé |
| --- | --- |
| Deux onglets et 24 alternances | Sélection native correcte, état initial restauré |
| Rafraîchissement partiel, masque/réapparition, désactivation/réactivation | Parcours exécutés ; captures réelles conservées |
| Peintures pendant les deux intervalles de repos | Compteur constant à 60 |
| Objets GDI avant/après les 24 alternances | 514 → 514 ; 512 au dernier relevé de repos |
| Sélection/défilement du code | Sélection relue, défilement vers la ligne 60, texte inchangé |
| Fermeture | Trace arrêtée puis classeur fermé sans sauvegarde, `Excel.Quit`, aucun hôte restant |
| Préférences et fichier de récupération de palette | SHA256 inchangés |

Les compteurs concernent le contrôle direct pour les peintures, et le processus Excel complet pour les objets GDI. Ces mesures ne prouvent pas l'absence de clignotement ailleurs. L'utilisateur a effectivement signalé un clignotement de **toute la barre d'outils** pendant ce passage.

## Trace graphique obtenue

La sonde C++ est indépendante du produit et reste sous `tools/probes/native-render-trace`. Elle observe les imports normaux de VBE7 en mémoire, sans changer les couleurs ni enregistrer le contenu du code. Son auto-test vérifie des pixels identiques avant/pendant/après, le transfert des appels et la restauration. Aucun binaire Microsoft n'est modifié sur disque.

Trace Excel réelle, VBE7 `7.01.1039` :

- **4 684 appels**, zéro événement perdu, zéro débordement du suivi des DC.
- **14 imports restaurés sur 14**, protections mémoire également restaurées.
- **107 sous-classes retirées sur 107**, arrêt avec statut zéro.
- Deux fenêtres Code identifiées : **1 075 `ExtTextOutA`**, dont **329 dessins de texte** et **746 remplissages sans texte**.
- **543 appels Code**, dont 182 dessins de texte, arrivent sans contexte de message ; leur HDC reste attribuable à une fenêtre Code. Une intervention limitée au seul `WM_PAINT` serait donc insuffisante.
- Propriétés : **400 appels sous `WM_DRAWITEM`**, dont 389 avec texte, et 52 remplissages sous `WM_ERASEBKGND`.
- 28 dessins de ComboBox utilisent un HDC mémoire dont la filiation n'est pas établie. Ils doivent rester exclus d'une recoloration fondée uniquement sur le contexte.

Les textes du code observés passent par la même adresse relative (`0x13DDF0`) pour cette version. Les catégories utilisent les couleurs natives grise, verte et cyan sur noir, et blanc sur bleu pour la sélection. Les adresses relatives sont des indices de diagnostic, jamais une API stable à coder en dur.

### Prochain moteur étayé par cette trace

Prototyper `ExtTextOutA` avec une liste explicite de HWND Code : attribuer le HDC à chaque appel, appliquer les couleurs avant le dessin, restaurer le HDC, traiter également les remplissages opaques, conserver les marqueurs et la sélection, laisser passer les DC non attribués. Les captures actuelles ne qualifient pas les breakpoints, tous les chemins de saisie, SOLIDWORKS ni les différents DPI.

## Signalement du clignotement des barres

Le code précédent programmait une passe sur tous les contrôles suivis après chaque peinture, notification ou changement géométrique local. Une notification Propriétés entraînait donc des écritures dans des barres indépendantes. Cette diffusion et les passes redondantes sont établies par le code.

La trace montre aussi 11 appels natifs vers la barre Standard avec texte noir et fond clair. Aucun n'a lieu pendant les 24 alternances mesurées ; ils ne suffisent donc pas à expliquer chaque clignotement observé. Les copies d'image réalisées par le complément et les dessins d'autres modules ne sont pas couverts par ces imports VBE7.

Une expérience a limité les mises à jour de titres à leur HWND et le rattrapage des barres aux activations/actions de code concernées. Dans `excel-toolbar-04`, les 24 alternances d'onglets n'ajoutent effectivement aucune passe sur les barres : compteurs de peinture 71 → 71, passes différées 62 → 62. Les objets GDI restent à 512 avant/après et le compteur des onglets reste à 60 au repos.

Cependant, les **153 captures exploitables** de la barre Standard pendant les changements de vue confirment le clignotement : les pixels clairs passent de **1,50 % à 35,00 %**. Plusieurs boutons et le champ de position deviennent clairs ; certains restent clairs dans la capture finale. La réduction des passes ne suffit donc pas et cette expérience n'est pas retenue par défaut. Deux essais précédents n'avaient aucune image exploitable à cause du contrôle de visibilité/masquage ; ils ne prouvent rien sur le rendu.

Le chemin historique de rattrapage est rétabli par défaut, avec exclusion des onglets Propriétés déjà dessinés directement. L'expérience reste accessible uniquement via `VBAi_NATIVE_LOCAL_REFRESH_EXPERIMENT=1` ; la sonde l'active avec `-LocalRefreshExperiment`, et force son absence pour un passage normal avant de restaurer l'environnement initial. La conversion existante évite maintenant la recopie du bitmap lorsqu'aucun pixel n'a changé. Le clignotement des barres **reste non résolu**.

La prochaine investigation doit couvrir les dessins Office des boutons, non observés par la seule trace des imports VBE7, afin de choisir les couleurs avant leur dessin. La fermeture des quatre instances Excel de cette campagne est passée par fermeture du classeur et `Excel.Quit`, sans arrêt forcé.

## Preuves et état de livraison

### Prototype des barres avant affichage — deuxième jalon

Le prototype est désormais implémenté dans `tools/probes/native-render-trace/ToolbarPatternPilot.h`, avec une entrée explicite `VBAiToolbarPatternStart`. Il ne modifie encore aucune dépendance de la DLL produit. Le mode de trace par défaut reste sans recoloration.

Trois chemins distincts ont été nécessaires :

1. Motifs `PatBlt/PATCOPY` : GDI reproduit le motif dans un tampon préalloué, avec son décalage natif. Les pixels neutres sont convertis, puis le résultat final est affiché. Les pixels précédents de l'écran ne sont jamais relus.
2. Icônes : les rectangles encore clairs passent par `SetDIBitsToDevice`, réellement observé ici en 24 bits BI_RGB, 16 × 16 pixels. Le prototype prépare une copie sombre des pixels neutres ou de la palette avant cet appel, en conservant les couleurs des glyphes et les données source.
3. Position ligne/colonne : VBE7 dessine le fond via `FillRect` et le texte via `ExtTextOutA`. Les couleurs sont remplacées temporairement avant le dessin et l'état du HDC est restauré.

Le premier essai `excel-pattern-pilot-01` traitait 36 motifs sans échec mais laissait les icônes et le champ clairs. Le deuxième traitait aussi les icônes et le texte, mais les captures étaient masquées par CorelDRAW et le fond du champ restait clair dans la capture native. Ces essais ne sont pas revendiqués comme une correction complète.

Le passage final `excel-pattern-pilot-03` comprend quatre alternances de vues Code/UserForm, avec échantillonnage pendant cinq secondes. La fenêtre jetable est temporairement maintenue visible sans activation ; l'observation rapporte **211 images exploitables en arrière-plan**, zéro changement d'activation et aucun masquage. Les captures minimale/maximale ne montrent pas de retour général au clair : pixels clairs entre **1,26 % et 3,44 %**, principalement glyphes/état du bouton sélectionné. Ce résultat reste une observation à la cadence mesurée, pas une preuve d'absence de tout flash plus bref.

Le journal confirme **26 motifs, 25 images d'icônes, 16 dessins de texte et 28 remplissages** traités, sans format refusé ni échec de motif sur ce parcours. Les **34 imports/protections et 111 sous-classes** sont restaurés, sans perte d'événement. Excel est fermé normalement, aucun hôte ne reste ouvert, préférences et récupération de palette inchangées. Le fichier `excel-pattern-pilot-03/verification.json` conserve les empreintes et compteurs.

L'auto-test ciblé vérifie 100 dessins de motifs avec origines et découpages variables, 100 remplissages, icônes 8/24/32 bits, sources et couleurs conservées, restitution des couleurs du texte et opérations non prises en charge intactes. Le compteur GDI revient à sa valeur initiale, une fois toutes les ressources de la fixture créées avant la mesure. Les premiers montages comptaient aussi des ressources créées par le test lui-même ; ils ont été corrigés. Aucune suite globale n'a été lancée.

**À intégrer et qualifier :** extraire ce moteur des outils de diagnostic vers une dépendance native du produit ; traiter la durée de vie de toutes les barres et fenêtres concernées ; remplacer les anciennes passes de conversion pour les surfaces effectivement couvertes ; vérifier popups, survols, états désactivés, plusieurs DPI, fermeture/réouverture, désactivation et SOLIDWORKS. La zone de code et les autres surfaces restent dans le périmètre global à finaliser.

### Investigation suivante : dessins effectifs de la barre

`excel-toolbar-trace-03` puis `excel-toolbar-trace-04` ciblent le module ayant enregistré la classe de la barre Standard, identifié par `GCLP_HMODULE` : **VBEUI.DLL**, distinct de VBE7.DLL. L'inspection des imports en mémoire est déterminante : les 14 fonctions graphiques sélectionnées passent ici par des imports différés déjà résolus. Les deux premières tentatives limitées aux imports ordinaires échouaient avec le statut 127, sans installer d'interception. L'inventaire sur disque ne remplace pas l'observation du module réellement chargé dans Office.

La sonde suit désormais ces imports différés sans forcer leur résolution et sans modifier leurs tables de déchargement. Son auto-test avec une fixture USER32 chargée à la demande confirme pixels identiques, état d'erreur conservé, rejet des autres threads, limite du journal, restauration des pointeurs/protections et maintien d'un import non résolu hors interception. Sources : [GCLP_HMODULE](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclasslongptrw), [imports PE](https://learn.microsoft.com/en-us/windows/win32/debug/pe-format).

Résultats du passage final `excel-toolbar-trace-04` :

- 234 appels enregistrés, zéro perte, zéro débordement des DC et arrêt de statut zéro.
- 14/14 imports et protections restaurés ; 111/111 sous-classes retirées.
- **39 appels PatBlt/PATCOPY utilisant un pinceau BS_PATTERN**, dont 36 vers la barre Standard. Ils ciblent directement le HDC de la barre, hors contexte WM_PAINT ; les rectangles correspondent aux boutons de 23 × 22 pixels et à leurs variantes.
- Les 11 ExtTextOutA observés dans ce module sont des remplissages sans texte, notamment pour les bordures et états sélectionnés. Le texte de position observé précédemment dans VBE7 est un autre chemin.
- 213 captures exploitables en arrière-plan, sans changement de focus imposé par la sonde : fraction de pixels clairs entre 1,26 % et 64,65 %. Le clignotement reste observable avec le chemin historique de rattrapage.
- Fermeture normale d'Excel, aucun hôte restant, empreintes des préférences et de la récupération de palette inchangées.

**Conséquence pour l'implémentation :** ne pas remplacer ces pinceaux par un aplat sombre. `BS_PATTERN` représente un motif bitmap ; ses pixels peuvent inclure les glyphes des boutons. La prochaine étape est d'inspecter ce motif et de préparer sa version sombre en préservant les icônes, puis de la sélectionner avant PatBlt pour les HWND explicitement reconnus. Le champ de position requiert aussi le traitement du chemin VBE7. Les pinceaux inconnus, les motifs non qualifiés et les fenêtres hors VBE doivent rester inchangés. Références : [LOGBRUSH](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logbrush), [PatBlt](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-patblt).

Cette investigation modifie uniquement les outils de diagnostic. La DLL produit installée conserve l'empreinte ci-dessous ; aucun correctif de pinceau n'est encore installé. Artefacts de compilation/auto-test : `trace-toolbar-final` ; rapport des empreintes et restaurations : `excel-toolbar-trace-04/verification.json`.

- Sources : `src/VBAi/Ui/VbeNativePropertyTabs.cs`, intégration dans `VbeNativeTheme.cs`.
- Tests et preuves : `artifacts/native-renderer-pilot/tests`, `excel-run-01`, `excel-toolbar-04`, `trace/selftest-result-final`.
- Synthèse de trace : `excel-run-01/native-render.summary.json` ; limites détaillées dans le README de la sonde.
- Versions antérieures de la DLL conservées dans `installed-baseline` et `installed-pilot-01` avant remplacement.
- Version de l'expérience barres conservée dans `installed-toolbar-04`. DLL finale installée : SHA256 `F18A26B6F538A525DE47BA50742096A9B12776FFC4C4B4902244A6B073DC95D9`, identique au résultat de compilation. `deployment-final.json` et `verification-final.json` confirment le remplacement, l'absence d'hôte restant et les préférences/palette inchangées.

La zone de code et les lignes Propriétés emploient encore le remappage d'image existant, y compris l'essai ClearType préparé avant la recherche. Leur netteté globale n'est pas déclarée corrigée. La sonde native de traçage n'est pas chargée automatiquement par l'add-in.

La compilation finale avec l'expérience locale désactivée par défaut réussit sans erreur ni avertissement. Ce dernier changement de sélection de chemin a été relu séparément ; il n'a pas donné lieu à un nouveau cycle Excel. Les validations réelles ci-dessus portent sur les versions indiquées dans leurs artefacts. La qualification complète SOLIDWORKS, le cycle complet de désactivation/réactivation et les différents DPI restent ouverts.

## Intégration dans l'add-in — 28 septembre 2026

Le moteur des barres est maintenant dans `src/VBAi.Native`, appelé par
`VbeNativeRenderer` lors de l'activation et de la désactivation du thème.
Le projet C# compile automatiquement la DLL native et l'embarque comme ressource.
Le chargement utilise un cache local nommé par SHA-256 ; l'empreinte et l'ABI sont
vérifiées avant activation. La solution Visual Studio présente les sources
natives dans un dossier dédié. La compilation nécessite la charge C++ x64
de Visual Studio ; les postes utilisateurs n'ont pas besoin du compilateur.

Le moteur ne traite que les barres enfants MsoCommandBar du VBE et leur thread.
Les menus flottants sont ignorés par ce moteur. Le rendu de secours existant
reste en place. Les allocations et la découverte des imports sont effectuées
hors des callbacks de dessin.

Contrôles exécutés sur cette intégration :
- Compilation C# net48/x64 et C++ /W4 /WX : aucune erreur ni avertissement.
- 20 cycles natifs activation/désactivation : interception limitée aux barres,
  rejet d'un arrêt depuis un autre thread, ajout et destruction de barres,
  exclusion d'un menu flottant, restauration des imports et équilibre GDI.
- Destruction de la racine : arrêt et retrait des références aux fenêtres.
- Extraction de la ressource embarquée, vérification d'empreinte et d'ABI,
  réutilisation du module chargé : succès dans un processus PowerShell isolé.

Ces tests ne prouvent pas la stabilité visuelle de l'intégration dans Office.
Le prototype a été observé sous Excel auparavant ; la version intégrée reste
à qualifier sous Excel puis SOLIDWORKS, avec les changements de visibilité,
les menus et les différentes résolutions/DPI. Une instance Excel utilisateur
était ouverte au terme de cette étape ; elle n'a pas été fermée ni modifiée.
La DLL installée n'a pas été remplacée.

Sorties de cette étape : `artifacts/native-product-integration`.

### Contrôles complémentaires et grandes barres — 28 septembre 2026

Le test de cycle de vie relève désormais indépendamment les adresses IAT et
leurs protections mémoire avant activation et les compare après chaque arrêt.
Le contrôle passe pour 20 cycles avec le moteur intégré.

Le tampon PATCOPY de 1024 × 128 pixels est réutilisé par portions pour les
surfaces plus grandes (limites explicites : 16384 × 2048). La phase de la brosse
est recalculée à chaque portion à partir du HDC de destination. Un test sur
2200 × 270 pixels compare chaque pixel au dessin natif suivi de la conversion
attendue, avec origines décalées et découpage traversant les raccords : succès.
Les tests existants des icônes, des couleurs de texte et de l'équilibre GDI
passent également. Ce résultat ne constitue pas une qualification multi-DPI.

Preuves : sorties compilées dans
`artifacts/native-product-integration/lifecycle-restoration` et
`artifacts/native-product-integration/wide-bars`.
La validation hôte attend toujours une session de test disponible.

## Validation Excel du moteur intégré — 28 septembre 2026, 16 h 51–16 h 55

Trois passages ciblés ont été exécutés, chacun dans une instance Excel jetable,
fermée normalement. Aucun processus Excel/SOLIDWORKS ne reste après les essais.

1. `excel-integrated-01` : 203 captures valides en arrière-plan sur Standard.
   Le moteur fonctionne mais les libellés de la barre de menus sont dégradés.
2. `excel-integrated-02` : 198 captures après correction de la conversion des
   couleurs déjà dans la palette sombre. Cette correction protège la rampe,
   mais ne suffit pas à résoudre la dégradation des libellés.
3. `excel-direct-01` : seconde passe de recoloration d'image désactivée pour les
   seules barres enfants prises en charge par le moteur. La comparaison visuelle
   confirme la disparition des rectangles clairs et la netteté retrouvée des
   libellés. Les icônes gardent leurs couleurs natives. 203 captures valides,
   toutes en arrière-plan, zéro erreur, zéro capture occultée ; proportion de
   pixels clairs entre 2,63 % et 4,31 %. Les extrêmes montrent les icônes et le
   champ de position, sans éclaircissement général de la barre.

Le dernier journal d'arrêt rapporte 118 portions PATCOPY, 135 images, 279
dessins de texte et 55 remplissages ; trois opérations non prises en charge,
aucun échec de dessin, neuf imports restaurés. Cela ne prouve ni l'absence de
clignotements entre captures, ni la couverture de tous les états natifs.

Le comportement validé en mode expérimental a été conservé par défaut : aucune
passe de recoloration d'image sur les barres enfants lorsque le moteur natif est
actif. Si son démarrage échoue, le secours existant reste disponible. La dernière
compilation (zéro erreur/avertissement) est installée :
`28FACD472388B945D62AD6A5790A3569AE11300D61E763C3C4C5315C116A10AB`.
La compilation finale retire seulement la condition expérimentale du chemin
testé ; elle n'a pas fait l'objet d'un quatrième passage Excel identique.

Les paramètres et le fichier de récupération de palette sont inchangés,
empreintes contrôlées dans `artifacts/native-product-integration/verification.json`.
L'ancienne DLL est conservée dans `installed-before` sous ce même dossier.

Restent notamment les menus flottants, les boutons système de la fenêtre fille,
la netteté du code/Propriétés, la zone de travail du concepteur, les changements
de visibilité, la désactivation réelle du thème, SOLIDWORKS et les DPI.
Le dessin utilisateur du UserForm lui-même ne doit pas être recoloré comme du chrome.

Observation hors correctif : le journal mentionne des icônes de menus de l'add-in
indisponibles (image source null). Cette anomalie n'a pas été corrigée ici ;
examiner le chargement des ressources d'icônes séparément.

## Observation passive sous SOLIDWORKS — 28 septembre 2026, 17 h 08

Session utilisateur conservée ouverte : SOLIDWORKS 2020 PID 498744, VBE sur
le module tets1. Aucun code modifié, aucune macro exécutée, aucune fermeture,
aucun changement de focus envoyé par la sonde. Le journal confirme le démarrage
du moteur natif (trois barres, dix imports initiaux) et la palette appliquée.

La barre Standard fournit 219 captures valides sur 5015 ms, toutes en arrière-plan,
sans changement d'activation, sans occultation ni erreur. La fraction de pixels
clairs reste à 2,18 %. La capture est sombre, avec icônes et champ de position
lisibles. Il s'agit d'un contrôle au repos : il ne qualifie pas les rafraîchissements
pendant saisie, survol, ouverture des menus ou redimensionnement.

La vue globale montre des libellés de menus nets, un fond de code cohérent avec
les panneaux et les onglets Propriétés sombres. Elle montre aussi des boutons
système de fenêtre fille et une barre de défilement des Variables locales encore
clairs. La finesse du texte du code reste à traiter via son rendu natif.
Les popups n'ont pas été ouverts lors de ce contrôle.

Les observations temporelles des deux autres barres ont été rejetées par le
contrôle d'occultation : zéro capture valide pour celles-ci. Ne pas les présenter
comme une preuve de stabilité. La capture Windows.Graphics.Capture a échoué ;
la capture globale provient du rectangle visible du VBE via la sonde passive.

Preuves locales :
C:/Users/jvc/Documents/Codex/2026-09-25/bo/artifacts/solidworks-native-observation
(observation.json, windows.json, bar-*.json, vbe-live.png, standard/*.png).
