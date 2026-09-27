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

## Persistance et portée

`%APPDATA%/CodexVBE/chat.db` conserve les sessions dans SQLite, via le runtime Windows `winsqlite3.dll`. Aucun serveur SQL ni paquet natif supplémentaire n'est nécessaire.

Une session possède un identifiant indépendant et une clé de document fondée sur le chemin complet du projet VBA enregistré. Ses messages, références, brouillon, fournisseur, modèle, effort, diffs et snapshots de restauration sont persistés. Deux documents différents n'utilisent pas la même liste de conversations. Un document sans chemin reçoit une portée temporaire propre à cette ouverture ; son association ne survit pas à un redémarrage.

Les chats Codex conservent aussi leur identifiant de thread. La reprise utilise `thread/resume` et retrouve le contexte serveur. Les autres fournisseurs conservent l'historique du protocole localement. Le changement de fournisseur démarre une conversation distincte.

La mémoire du document est une collection de notes éditables dans le panneau Chats. Elle est enregistrée localement dans une table séparée et n'est pas transmise par défaut. La case « Joindre les notes enregistrées au prochain message » ajoute explicitement ces notes au prochain envoi au fournisseur choisi. Le message conserve une section consultable des notes jointes.

Les clés API ne sont pas enregistrées dans cette base. Les contenus de conversation et les snapshots de code y sont stockés localement. En cas d'échec du stockage, l'interface indique que l'historique n'est pas enregistré.

## Vérification

Depuis le worktree :

```powershell
dotnet build CodexVBE.csproj -c Debug -p:OutputPath=artifacts/chat-build/final/ -p:BaseIntermediateOutputPath=artifacts/chat-build/obj/ -p:AppendTargetFrameworkToOutputPath=false
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatUx.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Conversation
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode History
powershell.exe -Sta -NoProfile -File tools/tests/Render-ChatUx.ps1 -Mode Reference -Width 460 -Height 850
```

Les tests utilisent un VBE simulé, une base SQLite temporaire réellement fermée puis rouverte, et des notifications du protocole injectées dans le client. Ils couvrent la navigation, le filtrage des procédures, les résumés, la non-duplication de la réponse finale, l'isolation des documents, plusieurs chats par document, le renommage, le brouillon, les références, les notes et les snapshots de rollback.

Les captures utilisent la vraie fenêtre WPF avec des données de démonstration sur le deuxième écran lorsqu'il est disponible. Elles ne constituent pas une validation dans Excel ou dans un hôte VBE réel. La reprise distante d'un thread authentifié et l'interruption d'une action COM en cours restent à valider dans l'hôte. Les fournisseurs autres que Codex affichent actuellement leur réponse finale, sans streaming ni résumé de réflexion.

## Références de conception

- [Copilot Visual Studio : contexte, références et historique](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-chat-context-references?view=visualstudio) : contexte explicite et navigation entre conversations.
- [Codex App Server](https://learn.chatgpt.com/docs/app-server) : événements de réponse, résumés de réflexion, reprise et interruption des tours.
- [Claude : artifacts](https://support.claude.com/en/articles/17153992-what-are-artifacts-and-how-do-i-use-them) : surfaces de travail consultables dans la conversation.
- [OpenClaw Control UI](https://docs.openclaw.ai/web/control-ui) : organisation des contrôles de conversation.

Cette interface n'annonce pas une parité complète avec ces produits. L'ancrage natif dans le VBE, les images jointes, les commandes slash, les modes Demander/Plan/Agent et la complétion dans l'éditeur restent des travaux distincts.
