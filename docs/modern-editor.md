# Éditeur Monaco

## Utilisation

Double-cliquer un module dans l’explorateur de projet du VBE ouvre son onglet Monaco. Les dossiers et les UserForms conservent leur comportement d’expansion et de conception. **Affichage → Éditeur VBAi** ouvre aussi la fenêtre. La fenêtre de code native reste derrière Monaco et ferme avec son onglet. Fermer Monaco conserve les brouillons puis ferme les fenêtres natives associées.

L’interface garde deux commandes permanentes : **Fermer le module** et **Ancrer/détacher**. La disposition fixe reste dans `ModernEditorWindow.Designer.cs`. Le WebView et les onglets de documents sont les éléments dynamiques. Le Designer ne démarre ni navigateur, ni worker, ni COM.

La synchronisation est continue après une pause de saisie. **Elle ne sauvegarde pas le classeur ou la macro sur disque** : enregistrer dans l’application hôte. Ctrl+S déclenche une synchronisation immédiate. Un conflit fait apparaître les actions de comparaison, rechargement et résolution. Une nouvelle modification native après comparaison interdit l’écrasement.

## Assistance VBA

- Suggestions des procédures, propriétés, variables, paramètres, constantes, types et champs déclarés dans le projet.
- Résolution des portées locales, des déclarations privées et des récepteurs typés simples (`objet.Membre`). Les brouillons ouverts remplacent les snapshots natifs dans l’index.
- Signatures, survol et navigation F12 vers les déclarations VBA, y compris les modules non ouverts.
- Lecture des membres et paramètres exposés par les bibliothèques COM référencées pour les types explicitement déclarés. Les interfaces héritées et l’interface par défaut d’une coclasse sont parcourues. Aucun objet métier n’est instancié pour cette analyse.
- Les résultats d’une requête pour une ancienne révision sont rejetés.

Ce résolveur couvre les déclarations et les récepteurs simples. Il ne remplace pas le compilateur : les chaînes d’appels complexes, `With`, la liaison tardive `Object/Variant`, les ambiguïtés entre bibliothèques homonymes et l’évaluation des branches conditionnelles ne sont pas résolues complètement. Les déclarations conditionnelles sont exclues des suggestions plutôt que présentées comme actives. Les noms de paramètres COM viennent des métadonnées ; leurs types/defaults ne sont pas reconstitués.

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

Le pas à pas nécessite le mode arrêt et utilise les commandes VBE identifiées, sans touches globales. La position affichée provient de la commande native **Afficher l’instruction suivante**, également demandée à l’entrée en mode arrêt ; ce n’est pas une lecture indépendante du pointeur d’exécution. Après un pas effectué directement dans le VBE, utiliser cette commande pour rafraîchir la position.

**Limite VBIDE : aucune collection publique de points d’arrêt n’est disponible.** Le cercle creux signifie donc « demande de bascule envoyée », avec une info-bulle demandant une vérification native. Il n’affirme pas qu’un point d’arrêt est installé. Les modifications du texte invalident ces marqueurs. Les concepteurs UserForm et le moteur de débogage restent natifs.

## Synchronisation, attributs et récupération

Un thread dédié prépare les différences, analyse les snapshots et chiffre les brouillons. Les accès au VBE et à WebView2 restent sur leur STA d’interface. Les callbacks WebView sont quittés avant de modifier les fenêtres WinForms. Chaque écriture vérifie l’identité COM, l’existence du composant, le mode conception, la protection et le texte précédent. La réconciliation conserve la révision capturée avant l’écriture pour ne pas écraser une frappe reçue pendant un appel COM.

Les caractères incompatibles avec la page de codes Windows sont refusés. Les écritures utilisent un patch de lignes et une relecture du formatage natif. En cas d’échec, un patch inverse est tenté ; tout échec de restauration est signalé.

Les corps des procédures avec attributs masqués peuvent être modifiés. Une édition qui remplace ou scinde leur déclaration est refusée, car elle peut supprimer les métadonnées cachées. Cette protection s’applique aussi au rollback. Le renommage ou remplacement de ces déclarations demande encore un workflow d’export/import préservant les attributs ; le remplacement automatique d’un composant vivant n’est pas effectué.

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

Le test `MonacoExcel`, activé par `VBAI_EDITOR_EXCEL_TEST=1`, utilise un classeur jetable : accents, synchronisation, conflit, sauvegarde/réouverture, renommage/suppression, récupération et conservation d’un attribut masqué. Le probe enregistré vérifie séparément le double-clic, l’ancrage/détachement et la fermeture couplée. Les réglages COM et AccessVBOM temporaires sont restaurés. L’arbre VBE est identifié par HWND et lu par MSAA : certains hôtes exposent `Window.HWnd=0` et aucun enfant UIA pour cet arbre.

Résultats détaillés locaux sous `artifacts/monaco/`. Les 32 Designers passent le chargement et le redimensionnement. Le parcours natif double-clic/ancrage/fermeture et le roundtrip Excel sont **PASS**. Les commandes de débogage/compilation utilisent les services natifs existants ; l’ensemble de leurs interactions visuelles dans Monaco n’a pas encore été qualifié dans tous les hôtes. SOLIDWORKS reste **NOT_RUN**, conformément à la demande.

## Sources techniques

- [API publique Monaco](https://github.com/microsoft/monaco-editor) : contrats de fournisseurs de langage.
- [Modèle d’objets VBIDE](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model) : code, volets et événements exposés.
- [Attributs VBA et export/import](https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations).
- [Distribution WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution).
