# Assistant LLM intégré au VBE

Cette première tranche ajoute une fenêtre de conversation WinForms au complément COM. Elle s'ouvre avec le VBE et tente d'être possédée par sa fenêtre principale via `MainWindow.HWnd`. L'assistant appelle directement `VbeSession`, sur le thread UI du VBE, sans serveur MCP, raccourci clavier ni coordonnées de fenêtre. Le menu natif **Affichage > Assistant CodexVBE** permet de rouvrir cette fenêtre ; **Outils > Configuration CodexVBE** ouvre les réglages.

## Fournisseurs

| Choix dans la fenêtre | État | Configuration avant lancement d'Excel ou SOLIDWORKS |
| --- | --- | --- |
| Codex (par défaut) | Implémenté via `codex app-server` local et l'authentification ChatGPT déjà gérée par le CLI | État ChatGPT et bouton de connexion dans Configuration ; aucune clé API requise. |
| OpenAI API | Implémenté via Chat Completions | Clé et URL dans Configuration ; `OPENAI_API_KEY` et `CODEXVBE_OPENAI_ENDPOINT` restent des valeurs de secours. |
| Ollama | Implémenté via son endpoint compatible Chat Completions | URL dans Configuration ; défaut `http://localhost:11434/v1/chat/completions`. |
| Claude | À venir | L'API et l'authentification Claude ne sont pas implémentées. |
| GitHub Copilot | À venir | L'authentification et l'API Copilot ne sont pas implémentées. |
| Gemini | À venir | L'API et l'authentification Gemini ne sont pas implémentées. |

Les URL distantes doivent utiliser HTTPS ; seul HTTP sur une adresse de boucle locale est accepté. Les réglages non secrets sont enregistrés pour l'utilisateur Windows dans `%APPDATA%\CodexVBE\settings.json`. Une éventuelle clé OpenAI API y est stockée uniquement sous forme chiffrée DPAPI `CurrentUser`, jamais dans le dépôt. Codex ne lit, ne copie et ne stocke aucun jeton OAuth : seul le processus `codex app-server` gère la connexion ChatGPT existante. La liste des modèles vient du fournisseur actif (`model/list`, `/v1/models` ou `/api/tags`) ; le choix persiste par fournisseur et s'applique au prochain message. Codex affiche aussi les niveaux de raisonnement annoncés pour le modèle choisi ; le niveau retenu est transmis comme `effort` au prochain `turn/start` et mémorisé par modèle. Changer de fournisseur efface l'historique affiché et le contexte transmis au modèle. Un modèle Ollama doit lui-même prendre en charge les appels d'outils pour piloter le VBE. Le catalogue OpenAI peut inclure des modèles qui ne prennent pas en charge Chat Completions ou les outils ; l'API signale ce cas lors de l'envoi.

Le client Codex initialise app-server par JSONL sur stdin/stdout, vérifie `account/read` avec `type: chatgpt`, crée un thread éphémère en mode lecture seule, puis transmet les commandes VBE comme `dynamicTools` (API expérimentale). Les appels `item/tool/call` retournent les résultats de `VbeSession` ; les demandes d'édition passent par la même validation visuelle que pour les autres fournisseurs. Un éventuel appel serveur inconnu reçoit une erreur explicite. Le choix de Codex ne passe pas par l'API OpenAI payante.

## Opérations proposées au modèle

Lectures : `list_projects`, `list_modules`, `read_module`, `list_forms`, `form_state`, `form_properties`, `form_control_properties`. Cette dernière énumère les propriétés de conception d'un contrôle avec leur type, état lecture seule, valeur lisible et éventuelle erreur COM. `open_form` ouvre le concepteur. Modifications : `replace_lines`, `create_form`, `add_form_control`, `rename_form_control`, `set_form_control_caption`, `set_form_control_font`, `set_form_control_geometry`.

Chaque fonction possède un schéma d'arguments. Le complément rejette les fonctions inconnues, les paramètres inconnus et les paramètres obligatoires manquants. Chaque modification ouvre une boîte de validation montrant l'intégralité des arguments. L'utilisateur doit cliquer **Autoriser** pour exécuter l'opération. `replace_lines` vérifie en plus l'empreinte SHA-256 du module ; les modifications de formulaire vérifient la version du formulaire et le mode conception. Ces vérifications sont faites après validation, juste avant l'appel COM. Les résultats sont relus par `VbeSession` et renvoyés au modèle.

Quand l'utilisateur envoie une demande, son texte, les résultats des lectures et l'historique de la conversation sont transmis au fournisseur choisi. En particulier, `read_module` transmet le code VBA au fournisseur. Aucune demande n'est envoyée automatiquement à l'ouverture du VBE.

## Vérification et limites

Le build .NET Framework 4.8 x64 a réussi dans une sortie isolée du worktree. Le parsing des tableaux de réponses JSON a été vérifié avec `JavaScriptSerializer` sous Windows PowerShell 5.1. Un essai direct du CLI local `codex-cli 0.156.1` a confirmé la connexion ChatGPT et l'appel d'un outil dynamique `vbe_status`, suivi de `turn/completed`. Un smoke test du nouveau client .NET en processus GUI sans console a ensuite confirmé `initialize`, `account/read` et `thread/start` avec outils dynamiques. Le test a découvert puis corrigé un BOM UTF-8 émis par `Process.StandardInput` en .NET Framework. L'interface WinForms initiale a été affichée dans Excel ; les nouveaux menus, réglages et tours Codex restent à tester dans cet hôte. L'ouverture de la fenêtre est protégée : un échec de l'UI est journalisé sans couper la passerelle existante.

Cette fenêtre de conversation n'est pas encore un panneau ancré. Les entrées Affichage/Outils ne sont pas encore validées en conditions réelles ; le streaming, l'annulation et les autres fournisseurs restent à faire. La couverture des outils VBE reste progressive ; voir `docs/vbe-coverage.md`.

Références API : [Codex App Server](https://learn.chatgpt.com/docs/app-server), [OpenAI Chat Completions et appels de fonctions](https://platform.openai.com/docs/api-reference/chat), [Ollama OpenAI compatibility](https://github.com/ollama/ollama/blob/main/docs/api/openai-compatibility.mdx).
