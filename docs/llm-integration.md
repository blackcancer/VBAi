# Assistant LLM intégré au VBE

Cette première tranche ajoute une fenêtre de conversation WinForms au complément COM. Elle s'ouvre avec le VBE et tente d'être possédée par sa fenêtre principale via `MainWindow.HWnd`. L'assistant appelle directement `VbeSession`, sur le thread UI du VBE, sans serveur MCP, raccourci clavier ni coordonnées de fenêtre.

## Fournisseurs

| Choix dans la fenêtre | État | Configuration avant lancement d'Excel ou SOLIDWORKS |
| --- | --- | --- |
| OpenAI API | Implémenté via Chat Completions | `OPENAI_API_KEY`, `CODEXVBE_OPENAI_MODEL`, facultatif `CODEXVBE_OPENAI_ENDPOINT` |
| Ollama | Implémenté via son endpoint compatible Chat Completions | `CODEXVBE_OLLAMA_MODEL`, facultatif `CODEXVBE_OLLAMA_ENDPOINT` ; défaut `http://localhost:11434/v1/chat/completions` |
| Codex | À venir | L'authentification Codex n'est pas implémentée. |
| Claude | À venir | L'API et l'authentification Claude ne sont pas implémentées. |
| GitHub Copilot | À venir | L'authentification et l'API Copilot ne sont pas implémentées. |
| Gemini | À venir | L'API et l'authentification Gemini ne sont pas implémentées. |

Les URL distantes doivent utiliser HTTPS ; seul HTTP sur une adresse de boucle locale est accepté. Les clés sont lues dans l'environnement du processus hôte, jamais écrites dans le dépôt. Changer de fournisseur efface l'historique affiché et le contexte transmis au modèle. Un modèle Ollama doit lui-même prendre en charge les appels d'outils pour piloter le VBE.

## Opérations proposées au modèle

Lectures : `list_projects`, `list_modules`, `read_module`, `list_forms`, `form_state`, `form_properties`. `open_form` ouvre le concepteur. Modifications : `replace_lines`, `create_form`, `add_form_control`, `rename_form_control`, `set_form_control_caption`, `set_form_control_font`, `set_form_control_geometry`.

Chaque fonction possède un schéma d'arguments. Le complément rejette les fonctions inconnues, les paramètres inconnus et les paramètres obligatoires manquants. Chaque modification ouvre une boîte de validation montrant l'intégralité des arguments. L'utilisateur doit cliquer **Autoriser** pour exécuter l'opération. `replace_lines` vérifie en plus l'empreinte SHA-256 du module ; les modifications de formulaire vérifient la version du formulaire et le mode conception. Ces vérifications sont faites après validation, juste avant l'appel COM. Les résultats sont relus par `VbeSession` et renvoyés au modèle.

Quand l'utilisateur envoie une demande, son texte, les résultats des lectures et l'historique de la conversation sont transmis au fournisseur choisi. En particulier, `read_module` transmet le code VBA au fournisseur. Aucune demande n'est envoyée automatiquement à l'ouverture du VBE.

## Vérification et limites

Le build .NET Framework 4.8 x64 a réussi dans une sortie isolée du worktree. Le parsing des tableaux de réponses JSON a été vérifié avec `JavaScriptSerializer` sous Windows PowerShell 5.1. La nouvelle interface et les appels réseau n'ont pas encore été exécutés dans Excel ou SOLIDWORKS ; la DLL actuellement chargée par Excel reste intacte. L'ouverture de la fenêtre est protégée : un échec de l'UI est journalisé sans couper la passerelle existante.

Cette fenêtre de conversation est un premier point d'entrée. Elle n'offre pas encore de bouton permanent dans la barre de commandes du VBE, ni de panneau ancré, streaming, annulation ou prise en charge des autres fournisseurs. La couverture des outils VBE reste progressive ; voir `docs/vbe-coverage.md` après intégration des branches.

Références API : [OpenAI Chat Completions et appels de fonctions](https://platform.openai.com/docs/api-reference/chat), [Ollama OpenAI compatibility](https://github.com/ollama/ollama/blob/main/docs/api/openai-compatibility.mdx).
