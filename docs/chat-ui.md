# Interface de conversation et sessions VBA

La structure fixe du chat est construite dans ChatWindow.Designer.cs et reste éditable avec le concepteur WinForms de Visual Studio. Le fil riche et le moteur de saisie avec correction orthographique utilisent des hôtes WPF fixes. Les puces de contexte, aperçus et diff disposent de vues WinForms indépendantes dans le Designer ; seules leurs instances et leurs données varient. Voir [les concepteurs WinForms](winforms-designer.md). Les contrôles WinForms personnalisés conservent les boutons arrondis et la carte de saisie. Le constructeur sans paramètre initialise uniquement le designer, sans ouvrir de session ni de base SQLite. Le sélecteur de document détermine le projet auquel appartiennent les conversations. Les réglages sont regroupés sous le champ de saisie ; le menu supérieur donne accès à la configuration et au rafraîchissement des modèles.

## Interactions disponibles

- Entrée envoie ; Maj+Entrée ajoute une ligne. Entrée ou Tab accepte d'abord une suggestion ouverte. La correction orthographique française reste active.
- `#` recherche les projets et modules ; `@` recherche les Sub, Function et Property Get/Let/Set. La liste est filtrée au curseur. Les puces montrent le contexte joint et permettent de le retirer.
- Les références sélectionnées sont cliquables dans les messages et les puces. La navigation relit le module et sélectionne la cible dans le VBE. Un projet ouvre la recherche de ses modules.
- Les réponses Codex arrivent progressivement. Les résumés de réflexion fournis par le serveur sont dépliables dans le fil, tout comme l'activité des outils. Les événements de raisonnement brut ne sont pas affichés.
- Avec la politique Automatique, les éditions s’appliquent sans dialogue répétitif, puis apparaissent sous forme de diff inline avec numéros de lignes, couleurs et bouton d’annulation. Le contrôle SHA empêche de remplacer des changements plus récents. Lecture seule interdit l’écriture ; Demander à chaque action conserve la validation configurée.
- Le bouton Modifications permet de rejoindre une carte de diff dans la conversation. Il n'ouvre pas de fenêtre d'approbation.
- Le bouton Arrêter interrompt le tour Codex ou la requête HTTP en cours. Une action VBE déjà exécutée reste dans l'historique et conserve son rollback.
- Le défilement suit les nouveaux messages tant que l'utilisateur reste en bas ; Dernier message ramène au fil actif.
- Nouveau ou Ctrl+N crée un chat. Le panneau Chats permet de chercher, reprendre, renommer, archiver et réactiver les conversations du document.

## Parcours de travail enrichis

- **Discussion / Plan / Agent** : le sélecteur est conservé par conversation. Discussion et Plan autorisent l’inspection et la compilation, mais bloquent les outils de modification et d’exécution dans les chemins synchrones et asynchrones. Le mode est transmis au fournisseur à chaque demande. Agent respecte toujours la politique VBE configurée. Le changement de mode est désactivé pendant une réponse.
- **Commandes `/`** : `/expliquer`, `/corriger`, `/refactoriser`, `/documenter`, `/tests`, `/plan`. Entrée ou Tab sélectionne une commande, puis ses paramètres restent éditables avant envoi. Accepter une suggestion choisit son mode ; saisir manuellement une commande conserve le mode visible. `/tests` demande de créer les tests et ne les exécute pas automatiquement.
- **Contexte** : « Joindre la sélection » capture les lignes et colonnes du panneau de code actif, dans le document de la conversation. Sans sélection étendue, la ligne courante est utilisée. Le SHA est revérifié avant envoi. Chaque sélection peut être retirée. Le panneau dépliable montre les références résolues, la sélection, les notes opt-in et leur taille en caractères. Le message conserve les snapshots effectivement transmis. La taille affichée concerne le contexte explicite, pas l’historique serveur ni les instructions système.
- **Historique** : recherche dans les titres, messages et snapshots de code ; épinglage ; export Markdown local ; « Créer une branche » depuis un message. Une branche copie le préfixe visible dans le même document et crée son propre contexte fournisseur, sans réutiliser le thread Codex ni dupliquer les droits de rollback des anciennes éditions.
- **Rollback** : diff découpé en blocs ; annulation d’un bloc, d’une modification ou de toute l’intervention. Une prélecture prépare tous les modules et détecte les conflits avant la première écriture. Chaque écriture vérifie encore le SHA vivant. Un bloc déplacé est retrouvé seulement si son contenu et jusqu’à trois lignes de contexte correspondent de façon unique. Les modifications indépendantes sont conservées. Une erreur lors des écritures suivantes indique le nombre de modules déjà restaurés : ce n’est pas une transaction COM atomique.
- **Vérification VBA** : bouton manuel et compilation après une intervention (case désactivable). Le résultat apparaît dans le chat. Une erreur localisable propose un lien vers le code et prépare une demande de correction. La correction reste un message à envoyer ; aucune macro ni batterie de tests n’est lancée automatiquement. « Aucun diagnostic natif observé » ne constitue pas une preuve d’exécution réussie des macros.
- **Intégration VBE** : actions Expliquer/Corriger/Refactoriser dans le menu contextuel du code ; commande « Fenêtre ancrable / flottante » dans le menu du chat. Le contrôle COM `CodexVBE.ChatToolWindow` est créé par `VBIDE.Windows.CreateToolWindow` et s'ouvre au démarrage dans le VBE. L'installation enregistre sa classe pour l'utilisateur courant et son ProgID dans la vue machine 64 bits, nécessaire à sa résolution par le VBE testé. Le VBE restaure la disposition mémorisée sans rattachement forcé au cadre principal. Si la zone cliente native est plus petite que le minimum du chat, seul ce panneau est détaché et rétabli dans une fenêtre native flottante ancrable, dans la zone de travail de l’écran du VBE. Il peut ensuite être ancré manuellement ; une disposition utilisable est conservée. Voir [le diagnostic de placement](chat-persistence-investigation.md). Si la création échoue, le chat reste flottant et affiche le motif.

Les nouvelles propriétés sont dans le payload JSON SQLite existant : les anciennes sessions restent lisibles, avec le mode Agent par défaut.

## Persistance et portée

`%APPDATA%/CodexVBE/chat.db` conserve les sessions dans SQLite, via le runtime Windows `winsqlite3.dll`. Aucun serveur SQL ni paquet natif supplémentaire n'est nécessaire.

Une session possède un identifiant indépendant et une clé de document fondée sur le chemin complet du projet VBA enregistré. Ses messages, références, brouillon, fournisseur, modèle, effort, diffs et snapshots de restauration sont persistés. Deux documents différents n'utilisent pas la même liste de conversations. Un document sans chemin reçoit une portée temporaire propre à cette ouverture ; son association ne survit pas à un redémarrage.

`%APPDATA%/CodexVBE/settings.json` conserve aussi le fournisseur par défaut, le dernier modèle choisi pour chaque fournisseur et l'effort choisi pour chaque modèle. Une nouvelle conversation prend ces valeurs par défaut. Reprendre une conversation restaure ses propres choix sans modifier les valeurs par défaut. Au démarrage, le fournisseur et son catalogue sont chargés même si Excel n'a pas encore exposé le projet VBA ; celui-ci est recherché de nouveau jusqu'à son apparition.

Les chats Codex conservent aussi leur identifiant de thread. La reprise utilise `thread/resume` et retrouve le contexte serveur. Les autres fournisseurs conservent l'historique du protocole localement. Le changement de fournisseur démarre une conversation distincte.

La mémoire du document est une collection de notes éditables dans le panneau Chats. Elle est enregistrée localement dans une table séparée et n'est pas transmise par défaut. La case « Joindre les notes enregistrées au prochain message » ajoute explicitement ces notes au prochain envoi au fournisseur choisi. Le message conserve une section consultable des notes jointes.

Les clés API ne sont pas enregistrées dans cette base. Les contenus de conversation et les snapshots de code y sont stockés localement. En cas d'échec du stockage, l'interface indique que l'historique n'est pas enregistré.

## Vérification

Les PR #2 et #3 sont intégrées. La mesure globale du code `fb166a4` atteint **100 % lignes et branches**, avec **1 151 tests réussis**, aucun échec et un scénario SOLIDWORKS non exécuté. Les deux essais Excel passent. Les 24 concepteurs WinForms ont été validés précédemment. Voir [le bilan courant](test-coverage.md). Les captures et essais d’interface ci-dessous sont des validations antérieures datées, avec leur propre périmètre.

Depuis le worktree :

```powershell
dotnet build src/CodexVBE/CodexVBE.csproj -c Debug -p:Platform=x64 -p:BuildOutputRoot="$PWD/artifacts/chat-build"
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatDesigner.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatUx.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatWorkflow.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Conversation
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode History
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Reference -Width 460 -Height 850
```

Les tests utilisent un VBE simulé, une base SQLite temporaire réellement fermée puis rouverte, et des notifications du protocole injectées dans le client. Ils couvrent la navigation, le filtrage des procédures, les résumés, la non-duplication de la réponse finale, l'isolation des documents, plusieurs chats par document, le renommage, le brouillon, les références, les notes et les snapshots de rollback.

Les tests Workflow couvrent en plus les refus d’outils synchrones/asynchrones en Discussion/Plan, l’interdiction d’écrire dans un autre projet, l’annulation par bloc et par intervention, les conflits, le code déplacé, la conservation d’éditions indépendantes, la recherche et l’export. Les tests UI vérifient la persistance du mode, de l’épinglage et des sélections, ainsi que la création d’une branche sans thread serveur partagé.

Validation locale du 27 septembre 2026 : compilation sans erreur ni avertissement ; suites ChatUx et ChatWorkflow passantes. Dans Excel, le chat a affiché le projet VBA, le fournisseur Codex, sept modèles du catalogue, le modèle par défaut et son effort. Les deux fenêtres WinForms s'ouvrent dans le concepteur Visual Studio sans erreur. Avec le ProgID machine enregistré, `VBIDE.Windows.CreateToolWindow` héberge le chat. Le premier ancrage automatique du VBE était en bas ; après déplacement manuel à droite, le VBE a conservé ce côté au redémarrage. Une nouvelle compilation a ouvert automatiquement le chat à droite dans Excel PID 41812 ; le contenu remplit le volet natif. Cette validation ne couvre pas encore SOLIDWORKS ni le premier placement sur une nouvelle installation. Menus contextuels, capture native de sélection et compilation avec localisation d'erreur : **NOT_RUN** dans cet essai.

Les captures modern-*.png composent le formulaire WinForms et ses deux zones WPF avec des données de démonstration sur le deuxième écran lorsqu'il est disponible. Elles ne constituent pas une validation dans Excel ou dans un hôte VBE réel. La reprise distante d'un thread authentifié et l'interruption d'une action COM en cours restent à valider dans l'hôte. Les fournisseurs compatibles et Claude affichent le texte progressivement par SSE ; Copilot utilise ses notifications. Bedrock affiche la réponse complète de Converse. Les résumés de réflexion progressifs restent propres à Codex. Voir providers.md pour les validations et limites.

## Références de conception

L’entrée **GitHub · synchroniser le VBA…** du menu du chat ouvre les vues Modifications Git et Historique,
avec comparaison côte à côte et commandes commit/push/fetch/pull séparées. Le dépôt reste dans un cache privé
sans fichiers adjacents à la macro. Voir [l’intégration GitHub](github-integration.md) pour le format VBA,
les sauvegardes, la restauration et les limites de validation.

- [Copilot Visual Studio : contexte, références et historique](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-chat-context-references?view=visualstudio) : contexte explicite et navigation entre conversations.
- [Codex App Server](https://learn.chatgpt.com/docs/app-server) : événements de réponse, résumés de réflexion, reprise et interruption des tours.
- [Claude : artifacts](https://support.claude.com/en/articles/17153992-what-are-artifacts-and-how-do-i-use-them) : surfaces de travail consultables dans la conversation.
- [OpenClaw Control UI](https://docs.openclaw.ai/web/control-ui) : organisation des contrôles de conversation.
- [VBIDE CreateToolWindow](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtoolwindow-method) : contrat de création de la fenêtre native.

Cette interface n’annonce pas une parité complète avec ces produits. Les images jointes et la complétion dans l’éditeur restent des travaux distincts. L’hébergement natif est présent ; le premier placement à droite sur une disposition vierge et les variantes d’hôtes/DPI restent à qualifier.

## Configuration, contexte et approbations

**Outils → Configuration VBAi…** ouvre les onglets Fournisseur, Compte GitHub et Apparence, même lorsque le chat est fermé. Les champs sont contextuels au fournisseur ; le choix des modèles et du raisonnement reste dans le chat. **Affichage → Assistant VBAi** rouvre le panneau.

La configuration propose Lecture seule, Demander à chaque action et Automatique. Les nouveaux paramètres ont Automatique comme valeur initiale ; une migration d’anciens paramètres sans politique explicite utilise Demander à chaque action. Les gardes de mode et de révision restent actives dans tous les cas. Discussion et Plan autorisent les inspections et la compilation, mais refusent les actions d’édition et d’exécution.

Le contexte système décrit l’hôte VBE, les projets, le mode, la sélection et les contraintes d’encodage relevées. Les notifications des projets, composants et références rafraîchissent le contexte ; les identités sont encore relues avant les actions. Une conversation dont le projet est fermé ou ambigu refuse les actions correspondantes.

Un chemin de fichier fourni par l’utilisateur peut être lu par un outil dédié, avec confirmation avant transmission au fournisseur. Cette lecture ne remplace pas le code vivant du VBE et ne dépend pas du niveau de raisonnement.

Les coupes de contrôles peuvent produire une carte de récupération dans le chat. La capture est conservée dans la session native ; elle ne doit pas être présentée comme récupérable après redémarrage sur la seule persistance du message.
