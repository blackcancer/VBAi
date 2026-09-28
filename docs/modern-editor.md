# Éditeur Monaco

## Utilisation

Double-cliquer un module dans l’explorateur de projet du VBE ouvre son onglet Monaco. Les dossiers et les UserForms conservent leur comportement d’expansion et de conception. **Affichage → Éditeur VBAi** ouvre aussi la fenêtre. La fenêtre de code native reste derrière Monaco et ferme avec son onglet. La croix de chaque onglet ferme le module associé et conserve son brouillon éventuel. La surface Monaco reste liée au VBE jusqu’à la fermeture de l’hôte.

Monaco occupe toute la zone centrale des documents VBE, sans bordure, commandes de fenêtre ni bouton ancrer/détacher. Les Designers UserForm et l’explorateur d’objets gardent leur place native ; sélectionner leur panneau Propriétés ne réaffiche pas Monaco par-dessus. La disposition fixe reste dans `ModernEditorWindow.Designer.cs`. Le WebView et les onglets de documents sont les éléments dynamiques. Le Designer ne démarre ni navigateur, ni worker, ni COM.

La synchronisation est continue après une pause de saisie. **Elle ne sauvegarde pas le classeur ou la macro sur disque** : enregistrer dans l’application hôte. Ctrl+S déclenche une synchronisation immédiate. Un conflit fait apparaître les actions de comparaison, rechargement et résolution. Une nouvelle modification native après comparaison interdit l’écrasement.

## Assistance VBA

- Suggestions des procédures, propriétés, variables, paramètres, constantes, types et champs déclarés dans le projet.
- Résolution des portées locales, des déclarations privées et des récepteurs typés et chaînes de propriétés/appels (`objet.Methode(...).Membre`), y compris les blocs `With` imbriqués. Les brouillons ouverts remplacent les snapshots natifs dans l’index.
- Signatures, survol et navigation F12 vers les déclarations VBA, y compris les modules non ouverts.
- Lecture des membres et paramètres exposés par les bibliothèques COM référencées pour les types explicitement déclarés. Les interfaces héritées et l’interface par défaut d’une coclasse sont parcourues. Aucun objet métier n’est instancié pour cette analyse.
- Les résultats d’une requête pour une ancienne révision sont rejetés.

Ce résolveur suit les types de retour COM sur six niveaux et conserve leur bibliothèque pour distinguer les types explicitement qualifiés. Il ne remplace pas le compilateur : la liaison tardive `Object/Variant`, les expressions arbitraires et l’évaluation des branches conditionnelles restent limitées. Les déclarations conditionnelles sont exclues des suggestions plutôt que présentées comme actives. Les noms de paramètres COM viennent des métadonnées ; leurs types/defaults ne sont pas reconstitués.

## Actions IA

Le menu contextuel et la palette Monaco proposent **Expliquer**, **Corriger**, **Refactoriser**. Les menus VBE correspondants utilisent aussi le document Monaco lorsqu’il est visible. La sélection, ou le module entier sans sélection, est attachée au chat du projet concerné ; aucune requête fournisseur n’est envoyée par le simple clic. Expliquer prépare le mode discussion ; corriger/refactoriser préparent le mode agent. Le lien de la pièce jointe revient au document Monaco. Un changement du brouillon invalide la pièce jointe avant envoi.

## Compilation et débogage

Le menu contextuel et la palette de commandes proposent :

| Commande | Raccourci |
|---|---|
| Compiler le projet | Ctrl+Maj+B |
| Basculer un point d’arrêt | F9 ou clic dans la marge |
| Pas à pas détaillé | F8 |
| Pas à pas principal | Maj+F8 |
| Pas à pas sortant | Ctrl+Maj+F8 |
| Afficher l’instruction suivante | Palette/menu contextuel |

La compilation est explicite, après synchronisation. Son premier diagnostic natif est reporté dans Monaco à la sélection laissée par le compilateur. Les marqueurs sont invalidés à la prochaine modification. La compilation LLM refuse également un brouillon non synchronisé. Aucune macro n’est exécutée par la synchronisation ou la compilation.

Le pas à pas nécessite le mode arrêt et utilise les commandes VBE identifiées, sans touches globales. La position affichée provient de la commande native **Afficher l’instruction suivante**, également demandée à l’entrée en mode arrêt ; ce n’est pas une lecture indépendante du pointeur d’exécution. Un changement de sélection native en mode arrêt relance le suivi ; la réapplication du marqueur après une correction synchronisée ne déplace pas le curseur Monaco. Une observation native échouée est réessayée.

**Limite VBIDE : aucune collection publique de points d’arrêt n’est disponible.** Le cercle creux signifie donc « demande de bascule envoyée », avec une info-bulle demandant une vérification native. Il n'affirme pas qu'un point d'arrêt est installé. Deux demandes successives retirent le marqueur local ; cela reste une représentation des demandes et non un inventaire confirmé du VBE. Le contrôle natif 51 est ciblé explicitement, sans confusion avec l’effacement global. Un essai Excel confirme que CommandBarButton.State reste à zéro avant/après bascule : cette propriété ne constitue pas une preuve utilisable. Les modifications du texte invalident ces marqueurs. Les concepteurs UserForm et le moteur de débogage restent natifs.

## Synchronisation, attributs et récupération

Un thread dédié prépare les différences, analyse les snapshots et chiffre les brouillons. Les accès au VBE et à WebView2 restent sur leur STA d’interface. Les callbacks WebView sont quittés avant de modifier les fenêtres WinForms. Chaque écriture vérifie l’identité COM, l’existence du composant, le mode conception ou arrêt, la protection et le texte précédent. La réconciliation conserve la révision capturée avant l’écriture pour ne pas écraser une frappe reçue pendant un appel COM.

En mode arrêt, une correction d’une seule ligne de corps de procédure utilise `ReplaceLine` sans réinitialiser l’exécution. Les changements de déclaration, insertions et suppressions restent en brouillon jusqu’au retour en conception. Les erreurs de synchronisation sont conservées dans le statut.

Les caractères incompatibles avec la page de codes Windows sont refusés. Les écritures utilisent un patch de lignes et une relecture du formatage natif. En cas d’échec, un patch inverse est tenté ; tout échec de restauration est signalé.

Les corps des procédures avec attributs masqués restent éditables. Le renommage et les modifications de signature sur une ligne sont pris en charge dans les modules standards, classes, modules de document et UserForms : export de sauvegarde, réassociation des attributs, rechargement du code seul dans le même composant et vérification du texte et des métadonnées par nouvel export. Les données du Designer ne sont pas rechargées. Les essais Excel vérifient notamment le membre par défaut et VB_PredeclaredId d'une classe, l'identité d'une feuille et le bouton d'un UserForm.

Les déclarations multilignes portant des attributs utilisent un [remplacement contrôlé autorisé](monaco-attribute-replacement-proposal.md) pour les modules standards, classes et UserForms : export complet, import sous un nom temporaire unique, contrôle du code et des métadonnées, réutilisation du FRX original et vérification des propriétés Designer lisibles, puis remplacement et reconnexion au nouvel objet COM. L'original est relu avant son retrait. Une erreur déclenche sa restauration ; les exports restent disponibles si la restauration échoue. Un essai Excel avec une panne injectée après retrait vérifie ce chemin de récupération. Les contrôles tiers ou propriétés Designer non vérifiables bloquent le remplacement. Les modules de document, dont l'identité appartient à une feuille ou au classeur, restent exclus de ce remplacement.

Les associations ambiguës (branches conditionnelles, métadonnées de paramètres après changement de signature, suppression/ajout simultané de procédures) restent protégées. Les variables de module portant des attributs ne sont pas remplacées silencieusement.

Un rechargement du code peut perdre les points d'arrêt natifs et l'historique Undo ; aucun inventaire public ne permet de les restaurer exactement. Si la restauration échoue, l'export original est conservé et son chemin est signalé.

Les brouillons sont chiffrés avec DPAPI sous `%LocalAppData%\CodexVBE\EditorDrafts`, sans service externe. Le nettoyage quotidien supprime les anciennes versions de plus de 30 jours, en conservant toujours le dernier fichier par module et les fichiers appartenant à un processus vivant. Les répertoires/jonctions de réanalyse sont ignorés. Un projet jamais enregistré n’a pas d’identité durable entre redémarrages. Une interruption avant réception/sauvegarde de la dernière frappe peut encore la perdre.

## Distribution

Monaco **0.55.1**, ses ressources, son worker, ses traductions et licences sont embarqués dans `EditorAssets`, sans CDN. WebView2 bloque les navigations externes, les permissions, les téléchargements, les fenêtres secondaires et les objets hôtes. Les commandes sont traduites dans les treize catalogues VBAi, sans service de traduction externe. Les widgets Monaco utilisent les traductions officielles disponibles ; arabe et hindi restent en anglais dans ce moteur.

Le paquet de release inclut Monaco, WebView2, ses loaders et `VBAi.Updater.exe`. Le futur installeur dispose des points d’entrée suivants :

- `VBAi.Updater.exe --check-webview2` : contrôle sans interface, code 0 si présent, 1 si absent, 2 en cas d’erreur.
- `VBAi.Updater.exe --ensure-webview2` : fenêtre WinForms de progression ; si nécessaire, téléchargement du bootstrapper Microsoft officiel, validation Authenticode et du signataire Microsoft, installation puis vérification de présence. Code de sortie 0 en cas de réussite, 1 en cas d’échec. Fermer la fenêtre termine le processus.

Le bootstrapper adapte l’architecture et le niveau d’installation au contexte Windows. Son installation réelle n’est pas déclenchée par les tests sur ce poste déjà équipé. L’installeur global et sa signature de publication restent un chantier distinct ; seul le prérequis est fourni ici.

Reconstruction reproductible des ressources : `tools/Build-MonacoAssets.ps1`, ou `-Offline` avec archives déjà en cache. Les versions et empreintes SHA-512 sont verrouillées.

## Vérification

Les tests unitaires couvrent les révisions, conflits, plans devenus obsolètes, snapshots immuables, attributs, rétention et refus d’exécution d’un prérequis non signé. Les tests `MonacoRuntime` chargent le vrai WebView2 : édition, diff, récupération, index de langage, marqueurs et contrats LLM avec frappes concurrentes.

Le test `MonacoExcel`, activé par `VBAI_EDITOR_EXCEL_TEST=1`, utilise un classeur jetable : accents, synchronisation, conflit, sauvegarde/réouverture, renommage/suppression, récupération, conservation et renommage d'une procédure avec attribut masqué. Il valide aussi la compilation via Monaco, le marqueur de position native et le pas à pas sur la seule procédure jetable Debug.Print. Le probe de cycle de vie n'exécute aucune macro. Le probe enregistré vérifie séparément le double-clic, l’ancrage/détachement et la fermeture couplée. Les réglages COM et AccessVBOM temporaires sont restaurés. L’arbre VBE est identifié par HWND et lu par MSAA : certains hôtes exposent `Window.HWnd=0` et aucun enfant UIA pour cet arbre.

Résultats détaillés locaux sous `artifacts/monaco/`. Les 32 Designers passent le chargement et le redimensionnement. Le parcours natif double-clic/ancrage/fermeture et le roundtrip Excel sont **PASS**. Les commandes de débogage/compilation utilisent les services natifs existants ; le parcours Excel de compilation sans erreur, instruction suivante, pas à pas et sortie est testé. Un vrai diagnostic de compilation Excel est également validé : identifiant non déclaré, sélection de la ligne native, marqueur Monaco, correction et disparition du marqueur. La capture des dialogues est isolée au PID propriétaire du VBE. Les autres hôtes restent à qualifier. SOLIDWORKS reste **NOT_RUN**, conformément à la demande.

### Intégration dans main

Les résultats ci-dessus appartiennent à la qualification de la branche Monaco. Le passage d'intégration sur `main` conserve les outils IDE existants et les cinq outils Monaco : **204 outils LLM**. Les tests d'attributs et de rétention sont rattachés aux miroirs `EditorVbeModule` et `EditorDraftStore` ; **184 miroirs pour 246 fichiers de production**. Les **32 surfaces WinForms** passent le chargement et le redimensionnement dans `artifacts/pr7-integration/designers/`.

Le premier passage réel a révélé une exception WinForms lors d'une fermeture pendant la création de contrôles. La fermeture attend maintenant la fin de l'initialisation, des opérations de synchronisation et de la disposition des contrôles d'état ; les callbacks ne réactualisent plus l'interface après une demande de fermeture. Deux scénarios supplémentaires vérifient la fermeture durant l'initialisation et durant la création du bouton de conflit, ainsi que la conservation des brouillons sans écriture VBA. Les exceptions de boucle UI sont remontées à VSTest au lieu de laisser un dialogue JIT bloquant.

Le lot ciblé avant le dernier delta de la branche donne **93 réussis, 0 échec, 1 ignoré** (`artifacts/pr7-integration/contracts-final/integration.trx`). La qualification globale et sa mesure de couverture sont suivies dans [le bilan de couverture](test-coverage.md). Les scénarios Excel sont désactivés dans ce passage d'intégration pour préserver les essais concurrents de l'autre session ; ce passage ne répète donc pas la preuve native de la branche. Les builds d'intégration utilisent `BuildOutputRoot` pour conserver la DLL chargée dans Excel.

### Diagnostics et attributs après PR #9

L'intégration conserve les **204 outils LLM** et compte **190 miroirs pour 254 fichiers de production**. Les **32 surfaces WinForms** passent leur qualification (`artifacts/pr9-integration/designers/designers.json`). La construction corrigée produit **0 erreur, 0 avertissement**.

Le scénario Excel final est **PASS** (`artifacts/pr9-integration/native-final/excel.trx`) : diagnostic réel, marqueur, correction, pas à pas, attributs standards/classes/formulaires, restauration après retrait et conservation des changements concurrents. Le correctif de fusion protège aussi le nom du composant lors d'un refus : seul le nom temporaire créé par l'opération peut être restauré. Voir [la qualification native](reference/native-qualification.md) et [le remplacement contrôlé](monaco-attribute-replacement-proposal.md) pour les limites.

## Sources techniques

- [API publique Monaco](https://github.com/microsoft/monaco-editor) : contrats de fournisseurs de langage.
- [Modèle d’objets VBIDE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model) : code, volets et événements exposés.
- [Attributs VBA et export/import](https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations).
- [Distribution WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution).

- [Limites de l'API VBIDE concernant les points d'arrêt](https://rubberduckvba.blog/using-rubberduck/) : constat publié par le projet Rubberduck, cohérent avec les essais natifs de cette intégration.
- [Thread STA et réentrance WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/threading-model).

## Régressions interface (septembre 2026)

- Commandes de débogage et synchronisation sérialisées avant les callbacks WebView ; cible d’onglet capturée avant les attentes asynchrones.
- F9 traité par Monaco et par le routage clavier WinForms/WebView2 ; clic dans la marge ou les numéros de ligne.
- Tests JavaScript : `node tools/tests/Test-MonacoLanguage.mjs` (11 cas, chaînes, portées, With et bibliothèques homonymes).
- Tests Excel isolés : F9 réel, exécution jusqu’au point demandé, clic de marge Chromium, correction en arrêt, compilation et actions IA sans fournisseur.
- Hébergement vérifié dans un MDI WinForms réel. Le script `Test-RegisteredMonaco.ps1` est adapté au nouveau parent MDI et aux croix d’onglets ; sa qualification dans Excel/VBE enregistré est réussie : parent MDIClient, remplissage de la zone, redimensionnement natif, double-clic projet et fermeture de l’onglet (`artifacts/monaco/host-integration-resize/monaco-host.json`). Les tests SOLIDWORKS restent différés.

Qualification de ce lot : **205 tests .NET réussis, 0 échec, 0 ignoré** (`artifacts/monaco/tests/fixes-final.trx`) et **11 tests JavaScript réussis**. Ce résultat ne constitue pas une mesure de couverture globale du projet.

La build locale de qualification est déployée dans `artifacts/monaco/host-build` du worktree. Les clés COM utilisateur AddIn et ChatToolWindow pointent vers cette build ; les anciennes valeurs sont sauvegardées dans `registration-before.json`. Aucun binaire du checkout main n’est remplacé.
