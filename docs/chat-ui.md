# Interface de conversation et sessions VBA

Le chat utilise une surface WPF unique dans la fenêtre hôte WinForms. Le sélecteur de document détermine le projet auquel appartiennent les conversations. Les réglages sont regroupés sous le champ de saisie ; le menu supérieur donne accès à la configuration et au rafraîchissement des modèles.

## Interactions disponibles

- Entrée envoie ; Maj+Entrée ajoute une ligne. Entrée ou Tab accepte d'abord une suggestion ouverte. La correction orthographique française reste active.
- `#` recherche les projets et modules ; `@` recherche les Sub, Function et Property Get/Let/Set. La liste est filtrée au curseur. Les puces montrent le contexte joint et permettent de le retirer.
- Les références sélectionnées sont cliquables dans les messages et les puces. La navigation relit le module et sélectionne la cible dans le VBE. Un projet ouvre la recherche de ses modules.
- Les réponses Codex arrivent progressivement. Les résumés de réflexion fournis par le serveur sont dépliables dans le fil, tout comme l'activité des outils. Les événements de raisonnement brut ne sont pas affichés.
- Les éditions `replace_lines` s'appliquent sans dialogue, puis apparaissent sous forme de diff inline avec numéros de lignes, couleurs et bouton d'annulation. Le contrôle SHA empêche de remplacer des changements plus récents. Lecture seule interdit toujours l'écriture.
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
- **Intégration VBE** : actions Expliquer/Corriger/Refactoriser dans le menu contextuel du code ; commande « Fenêtre ancrable / flottante » dans le menu du chat. Cette commande tente de créer le contrôle COM `CodexVBE.ChatToolWindow` par `VBIDE.Windows.CreateToolWindow`. L’installation enregistre ce contrôle pour l’utilisateur courant. Si VBE refuse sa création, la fenêtre reste flottante et affiche le motif.

Les nouvelles propriétés sont dans le payload JSON SQLite existant : les anciennes sessions restent lisibles, avec le mode Agent par défaut.

## Persistance et portée

`%APPDATA%/CodexVBE/chat.db` conserve les sessions dans SQLite, via le runtime Windows `winsqlite3.dll`. Aucun serveur SQL ni paquet natif supplémentaire n'est nécessaire.

Une session possède un identifiant indépendant et une clé de document fondée sur le chemin complet du projet VBA enregistré. Ses messages, références, brouillon, fournisseur, modèle, effort, diffs et snapshots de restauration sont persistés. Deux documents différents n'utilisent pas la même liste de conversations. Un document sans chemin reçoit une portée temporaire propre à cette ouverture ; son association ne survit pas à un redémarrage.

`%APPDATA%/CodexVBE/settings.json` conserve aussi le fournisseur par défaut, le dernier modèle choisi pour chaque fournisseur et l'effort choisi pour chaque modèle. Une nouvelle conversation prend ces valeurs par défaut. Reprendre une conversation restaure ses propres choix sans modifier les valeurs par défaut. Au démarrage, le fournisseur et son catalogue sont chargés même si Excel n'a pas encore exposé le projet VBA ; celui-ci est recherché de nouveau jusqu'à son apparition.

Les chats Codex conservent aussi leur identifiant de thread. La reprise utilise `thread/resume` et retrouve le contexte serveur. Les autres fournisseurs conservent l'historique du protocole localement. Le changement de fournisseur démarre une conversation distincte.

La mémoire du document est une collection de notes éditables dans le panneau Chats. Elle est enregistrée localement dans une table séparée et n'est pas transmise par défaut. La case « Joindre les notes enregistrées au prochain message » ajoute explicitement ces notes au prochain envoi au fournisseur choisi. Le message conserve une section consultable des notes jointes.

Les clés API ne sont pas enregistrées dans cette base. Les contenus de conversation et les snapshots de code y sont stockés localement. En cas d'échec du stockage, l'interface indique que l'historique n'est pas enregistré.

## Vérification

Depuis le worktree :

```powershell
dotnet build CodexVBE.csproj -c Debug -p:OutputPath=artifacts/chat-build/final/ -p:BaseIntermediateOutputPath=artifacts/chat-build/obj/ -p:AppendTargetFrameworkToOutputPath=false
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatUx.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatWorkflow.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Conversation
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode History
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Reference -Width 460 -Height 850
```

Les tests utilisent un VBE simulé, une base SQLite temporaire réellement fermée puis rouverte, et des notifications du protocole injectées dans le client. Ils couvrent la navigation, le filtrage des procédures, les résumés, la non-duplication de la réponse finale, l'isolation des documents, plusieurs chats par document, le renommage, le brouillon, les références, les notes et les snapshots de rollback.

Les tests Workflow couvrent en plus les refus d’outils synchrones/asynchrones en Discussion/Plan, l’interdiction d’écrire dans un autre projet, l’annulation par bloc et par intervention, les conflits, le code déplacé, la conservation d’éditions indépendantes, la recherche et l’export. Les tests UI vérifient la persistance du mode, de l’épinglage et des sélections, ainsi que la création d’une branche sans thread serveur partagé.

Validation locale du 27 septembre 2026 : compilation sans erreur ni avertissement ; suites ChatUx et ChatWorkflow passantes. Dans Excel, le chat a affiché le projet VBA, le fournisseur Codex, sept modèles du catalogue, le modèle par défaut et son effort. Les deux fenêtres WinForms s'ouvrent dans le concepteur Visual Studio sans erreur. L'activation COM du contrôle du chat réussit dans la session utilisateur, mais `VBIDE.Windows.CreateToolWindow` refuse encore de l'héberger avec `CO_E_CLASSSTRING` ; l'ancrage est **FAILED**, la fenêtre flottante reste utilisable. Menus contextuels, capture native de sélection et compilation avec localisation d'erreur : **NOT_RUN** dans cet essai. Les captures `surface-*.png` rendent directement la surface WPF pour éviter les problèmes de cadrage ou de recouvrement entre écrans.

Les captures utilisent la vraie fenêtre WPF avec des données de démonstration sur le deuxième écran lorsqu'il est disponible. Elles ne constituent pas une validation dans Excel ou dans un hôte VBE réel. La reprise distante d'un thread authentifié et l'interruption d'une action COM en cours restent à valider dans l'hôte. Les fournisseurs autres que Codex affichent actuellement leur réponse finale, sans streaming ni résumé de réflexion.

## Références de conception

- [Copilot Visual Studio : contexte, références et historique](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-chat-context-references?view=visualstudio) : contexte explicite et navigation entre conversations.
- [Codex App Server](https://learn.chatgpt.com/docs/app-server) : événements de réponse, résumés de réflexion, reprise et interruption des tours.
- [Claude : artifacts](https://support.claude.com/en/articles/17153992-what-are-artifacts-and-how-do-i-use-them) : surfaces de travail consultables dans la conversation.
- [OpenClaw Control UI](https://docs.openclaw.ai/web/control-ui) : organisation des contrôles de conversation.
- [VBIDE CreateToolWindow](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtoolwindow-method) : contrat de création de la fenêtre native.

Cette interface n’annonce pas une parité complète avec ces produits. Les images jointes et la complétion dans l’éditeur restent des travaux distincts. L'hébergement natif du contrôle dans le VBE reste à corriger.
