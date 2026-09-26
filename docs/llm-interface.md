# Interface LLM du VBE

Référence d'interaction : GitHub Copilot Chat dans Visual Studio. Son sélecteur de modèles se trouve dans la fenêtre de conversation ; la gestion des modèles et des fournisseurs s'ouvre depuis ce sélecteur. CodexVBE conserve aussi une entrée `Outils > Configuration CodexVBE…`, demandée pour rendre ces réglages accessibles depuis VBE.

## Configuration du fournisseur

- La configuration affiche le fournisseur courant et seulement les champs qui lui sont propres. Changer de fournisseur remplace le contenu du formulaire ; les champs des autres fournisseurs ne restent pas visibles.
- **Codex** : état du compte ChatGPT utilisé par `codex app-server`, action de connexion si nécessaire. Aucune clé OpenAI API n'est demandée pour ce mode.
- **OpenAI API** : clé API et, si l'implémentation le permet, URL d'API personnalisée. L'absence de crédit API n'empêche pas d'utiliser Codex avec ChatGPT.
- **Ollama** : adresse de l'instance locale et état de connexion ; aucune clé API par défaut.
- **Claude, GitHub Copilot, Gemini** : afficher uniquement les paramètres nécessaires lorsque leur intégration est implémentée. Un fournisseur non implémenté doit être indiqué comme tel, sans promettre une connexion fonctionnelle.
- Les secrets sont conservés pour le compte Windows courant et ne sont jamais affichés après enregistrement. Une erreur de connexion doit rester visible dans la configuration.

## Choix du modèle dans la conversation

- Un sélecteur de modèles est présent dans la fenêtre de conversation, près de la zone de saisie comme dans Visual Studio. Il affiche les modèles réellement retournés pour le fournisseur et le compte configurés, ainsi qu'un état de chargement, une action d'actualisation et une erreur explicite si la liste ne peut pas être chargée.
- La liste dépend du fournisseur sélectionné : `model/list` pour Codex app-server ; `GET /v1/models` pour OpenAI API et Claude ; `GET /api/tags` pour Ollama ; `models.list` pour Gemini. Filtrer les modèles qui ne peuvent pas servir au chat ou aux appels d'outils lorsque l'API donne cette capacité. Les identifiants issus d'une liste publique statique ne prouvent pas l'accès du compte.
- Le choix du modèle persiste pour le fournisseur. Après actualisation, si ce modèle a disparu, le sélecteur revient au modèle par défaut annoncé par le fournisseur ou demande un nouveau choix ; il ne transmet pas silencieusement un identifiant périmé.
- Changer de modèle pendant qu'une réponse est en cours ne modifie pas cette réponse. Le nouveau modèle sert au tour suivant. Pour Codex, `turn/start.model` permet cette modification sur la conversation courante.
- Le niveau de raisonnement est choisi à côté du modèle dans la conversation. Pour Codex, les choix viennent de `supportedReasoningEfforts` du modèle sélectionné ; `defaultReasoningEffort` fournit la valeur initiale, et `turn/start.effort` transmet le choix au tour suivant. Masquer ou désactiver ce réglage lorsqu'un fournisseur ou un modèle ne le propose pas.

## Présentation et vérification visuelle

- La fenêtre de conversation garde un en-tête lisible, un historique qui occupe l'espace disponible et une zone de saisie compacte. Le fournisseur, le modèle, l'effort, l'actualisation et la configuration restent chacun identifiables et accessibles à une taille de fenêtre réduite et sous la mise à l'échelle Windows.
- La configuration ajuste sa hauteur au contenu du fournisseur : trois lignes Codex ne doivent pas laisser le grand vide d'un formulaire prévu pour tous les fournisseurs. Les actions principales restent près des champs. Les éléments masqués ne doivent pas réserver de place.
- Valider la disposition dans une capture réelle du VBE Excel après compilation ; la réussite du build et l'arbre d'accessibilité ne garantissent pas l'alignement visuel.

## Accès VBE

- `Affichage > Assistant CodexVBE` rouvre l'assistant après fermeture. `Outils > Configuration CodexVBE…` ouvre directement les paramètres du fournisseur. Ces commandes sont ajoutées aux menus natifs du VBE et cherchées par leur intitulé selon la langue de l'hôte.
- Le panneau latéral ancré est une étape distincte de l'interface flottante actuelle. L'essai Excel a validé `CreateToolWindow` et `LinkedWindows.Add` avec `Shell.Explorer.2`, mais pas encore l'hébergement du chat ni son positionnement à droite.

## Niveaux d'approbation des commandes VBE

- La configuration propose **Lecture seule** (les commandes de modification VBE sont refusées), **Demander à chaque action** (une validation avant chaque commande de modification) et **Automatique** (les commandes de modification VBE s'exécutent sans boîte répétée). Pour le prototype et les essais en cours, le niveau initial est **Automatique**, conformément à la préférence exprimée par l'utilisateur.
- Le niveau s'applique à toutes les écritures VBE, y compris la création d'un module standard, d'une classe ou d'un UserForm, le remplacement de lignes et la conception d'un contrôle. Une commande nouvelle doit être classée avant exposition au LLM ; elle ne doit pas contourner le réglage.
- Les commandes ciblent un projet et un composant explicites et relisent l'état après l'action. Les contrôles de mode Conception, de collision de nom et de révision restent actifs au niveau Automatique.
- La lecture d'un fichier fourni par chemin demeure une autorisation distincte avant sa transmission au fournisseur LLM ; elle n'est pas une écriture VBE. Le niveau de raisonnement du modèle ne modifie aucun droit.

## Contexte de l'agent VBE

- Les instructions de la conversation indiquent que l'agent est intégré au **Visual Basic Editor** du processus hôte courant. Le nom de l'hôte est relevé dans le processus (par exemple Excel ou SOLIDWORKS) ; les projets, composants, formulaires, sélections et modes sont relus à travers les commandes VBE. Ne jamais déduire le contenu de SOLIDWORKS d'un essai Excel.
- L'agent utilise les commandes VBE exposées par le complément pour lire et piloter **l'éditeur vivant**. Il peut aussi lire un fichier externe lorsque l'utilisateur en fournit le chemin ; le contexte système ne doit pas interdire les accès au système de fichiers nécessaires à cette demande. Cette lecture ne remplace pas la relecture de l'état vivant du VBE. Aucun raccourci clavier ni clic fondé sur des coordonnées n'est une commande produit. Il cible explicitement projet, module, formulaire et contrôle ; il lit l'état et la révision avant écriture et suit le niveau d'approbation choisi.
- Le contexte initial décrit les capacités effectivement implémentées et leurs limites. Il n'annonce pas une couverture intégrale de VBE tant que les références, la boîte à outils, les fenêtres de débogage et les autres surfaces ne sont pas validées. L'état vivant peut changer entre deux messages : l'agent le relit plutôt que d'utiliser un inventaire figé dans son prompt.

Sources : [Copilot Chat et modèles dans Visual Studio](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-usage-and-models?view=visualstudio), [sélecteur Visual Studio dans la documentation GitHub](https://docs.github.com/en/copilot/how-tos/use-ai-models/change-the-chat-model), [Codex app-server](https://learn.chatgpt.com/docs/app-server), [modèles OpenAI API](https://platform.openai.com/docs/api-reference/models/object), [modèles Ollama](https://docs.ollama.com/api/tags), [modèles Claude](https://platform.claude.com/docs/en/api/models/list), [modèles Gemini](https://ai.google.dev/api/models).
