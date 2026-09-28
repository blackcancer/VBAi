# Fournisseurs du chat

Le catalogue comporte Codex, OpenAI API, Ollama, Claude, GitHub Copilot, Gemini, Mistral, DeepSeek, OpenRouter, LM Studio, Personnalisé (OpenAI), Azure OpenAI, Grok, Groq et Amazon Bedrock. Leur disponibilité dans la liste signifie que le transport est implémenté ; elle ne prouve pas que le compte, le serveur ou le modèle sélectionné est disponible.

## Configuration

La fenêtre comporte trois onglets : **Fournisseur**, **Compte GitHub** et **Apparence**. Le compte GitHub des dépôts est indépendant de l’authentification du fournisseur IA. Les modèles et le niveau de raisonnement se choisissent dans la conversation, et non dans cette fenêtre.

Dans **Paramètres du fournisseur**, sélectionner le fournisseur et renseigner sa clé et, si nécessaire, son URL complète. Enregistrer, puis choisir un modèle dans le chat. Changer de fournisseur dans les paramètres conserve les brouillons séparément jusqu’à Enregistrer ; Annuler les abandonne. Une clé vide conserve la valeur enregistrée ; la case de suppression retire la clé enregistrée et réactive le repli éventuel sur la variable d’environnement.

| Fournisseur | URL par défaut | Variable de clé |
| --- | --- | --- |
| OpenAI API | `https://api.openai.com/v1/chat/completions` | `OPENAI_API_KEY` |
| Ollama | `http://localhost:11434/v1/chat/completions` | Clé facultative dans les paramètres |
| Claude | `https://api.anthropic.com/v1/messages` | `ANTHROPIC_API_KEY` |
| Gemini | `https://generativelanguage.googleapis.com/v1beta/openai/chat/completions` | `GEMINI_API_KEY` |
| Mistral | `https://api.mistral.ai/v1/chat/completions` | `MISTRAL_API_KEY` |
| DeepSeek | `https://api.deepseek.com/chat/completions` | `DEEPSEEK_API_KEY` |
| OpenRouter | `https://openrouter.ai/api/v1/chat/completions` | `OPENROUTER_API_KEY` |
| LM Studio | `http://localhost:1234/v1/chat/completions` | `LM_STUDIO_API_KEY` (facultative) |
| Personnalisé (OpenAI) | URL complète à renseigner | `CODEXVBE_CUSTOM_API_KEY` (facultative) |
| Azure OpenAI | `https://<ressource>.openai.azure.com/openai/v1/chat/completions` à renseigner | `AZURE_OPENAI_API_KEY` ou `AZURE_OPENAI_ENTRA_TOKEN` suivant le mode |
| Grok | `https://api.x.ai/v1/chat/completions` | `XAI_API_KEY` |
| Groq | `https://api.groq.com/openai/v1/chat/completions` | `GROQ_API_KEY` |
| Amazon Bedrock | `https://bedrock-runtime.<région>.amazonaws.com` à renseigner | `AWS_BEARER_TOKEN_BEDROCK` |

Les URL de secours suivent `CODEXVBE_<FOURNISSEUR>_ENDPOINT` avec les identifiants `OPENAI`, `OLLAMA`, `CLAUDE`, `GEMINI`, `MISTRAL`, `DEEPSEEK`, `OPENROUTER`, `LMSTUDIO`, `CUSTOM`, `AZURE`, `GROK`, `GROQ`, `BEDROCK`. Les réglages enregistrés ont priorité. Redémarrer l’hôte après une modification des variables d’environnement.

### Fournisseur personnalisé, Azure et Bedrock

Ces trois entrées utilisent une liste de modèles explicite, un identifiant par ligne dans Configuration. Cette liste peut aussi provenir de `CODEXVBE_CUSTOM_MODEL`, `CODEXVBE_AZURE_MODEL` ou `CODEXVBE_BEDROCK_MODEL`. Le catalogue ne dépend donc pas d’un endpoint `/models` que certains serveurs n’exposent pas.

Le fournisseur personnalisé permet de choisir un nom d’affichage et une URL compatible Chat Completions ; sa clé Bearer est facultative. Son identité persistée reste stable si son nom d’affichage change. Un seul profil personnalisé est actuellement proposé.

Azure utilise l’API v1 et les **noms de déploiement** saisis, plutôt qu’une liste de modèles théoriquement disponibles chez Azure. La clé passe dans `api-key`. La case Microsoft Entra utilise un jeton Bearer : fournir le jeton dans le champ de secret ou `AZURE_OPENAI_ENTRA_TOKEN`, puis le renouveler à expiration. Si une clé API était enregistrée avant le changement de mode, la remplacer par le jeton ou la supprimer pour utiliser la variable d’environnement. Cette version ne lance pas de connexion Entra interactive et ne renouvelle pas les jetons automatiquement.

Bedrock utilise l’API native **Converse**, l’URL Runtime de la région choisie et l’ID ou ARN du modèle/profil d’inférence. Les définitions d’outils, résultats et contenus signés sont convertis et conservés entre les tours. L’authentification implémentée est la **clé API Bedrock Bearer**, distincte d’une paire IAM access key/secret key ; la signature SigV4 et les profils AWS ne sont pas implémentés. La clé doit permettre d’invoquer le modèle ou profil choisi dans cette région.

Les clés sont chiffrées avec DPAPI pour le compte Windows courant. Elles restent dans les paramètres locaux, séparées par fournisseur ; elles ne sont pas enregistrées dans SQLite. Les anciennes propriétés OpenAI/Ollama restent lisibles. Les URL distantes exigent HTTPS ; HTTP est accepté uniquement sur une adresse de boucle locale. Les redirections HTTP automatiques sont désactivées. Les erreurs affichent le fournisseur et le statut HTTP, sans recopier le corps de la réponse susceptible de contenir des données sensibles.

## Isolation des conversations

Les processus CLI de VBAi ont leur propre stockage local :

- Codex : `%LOCALAPPDATA%\CodexVBE\Providers\Codex`, transmis comme `CODEX_HOME`.
- Copilot : `%LOCALAPPDATA%\CodexVBE\Providers\Copilot`, transmis comme `COPILOT_HOME`.

Ces variables sont fixées seulement dans les processus enfants, y compris les commandes de connexion et de statut. Les clients habituels conservent leurs dossiers par défaut. L’historique SQLite de VBAi reste séparé du stockage interne de chaque CLI. Une connexion dans les paramètres de VBAi peut être nécessaire dans le nouveau dossier ; aucun fichier d’authentification personnel n’est copié.

Une ancienne conversation Codex est reprise à partir de son historique local comme contexte, puis reçoit un nouveau fil privé. Son ancien identifiant externe n’est pas repris. Les fils déjà privés conservent leur reprise native. Cette migration ne supprime pas les anciens fils déjà inscrits dans l’historique externe.

Les fournisseurs HTTP (OpenAI API, Claude, Gemini, Ollama et autres) n’exécutent pas de CLI de conversation : leur historique local est conservé par VBAi. Cela ne modifie pas la conservation éventuelle de données côté service distant.

Les emplacements suivent la [configuration officielle Codex](https://developers.openai.com/codex/config-advanced/) et la [référence Copilot](https://docs.github.com/en/copilot/reference/copilot-cli-reference/cli-command-reference).

## Codex : fournisseur prioritaire

Codex est le fournisseur par défaut. Le complément utilise le processus `codex app-server` et le compte ChatGPT authentifié par le CLI, sans clé OpenAI API. La configuration propose l’état du compte, la connexion et l’actualisation ; les champs de clé et d’endpoint des transports HTTP sont masqués pour ce mode.

Le catalogue provient de `model/list`. Les niveaux de raisonnement disponibles et la valeur initiale proviennent des métadonnées du modèle. Le chat transmet le modèle et l’effort au prochain tour, conserve l’identifiant du thread par session et reprend celui-ci avec `thread/resume`.

Le changement de fournisseur actualise son catalogue. Les réglages par défaut et ceux de chaque conversation sont persistés séparément : voir [Conversation et sessions](chat-ui.md). Un abonnement ChatGPT et des crédits OpenAI API sont des moyens d’accès distincts.

## Claude et fournisseurs HTTP

Claude utilise l’API Messages : instructions système séparées, conversion des définitions en `input_schema`, des appels en `tool_use` et des résultats en `tool_result`. Les résultats parallèles sont regroupés dans un message utilisateur. Le catalogue suit la pagination Anthropic. La réponse est limitée à 8 192 jetons ; une réponse tronquée n’est pas exécutée comme une action complète.

Les autres fournisseurs HTTP utilisent Chat Completions. `store=false` est envoyé uniquement à OpenAI. Les métadonnées retournées dans les messages, notamment les signatures Gemini et le contenu de raisonnement requis par certains modèles, sont conservées pour les tours suivants. Le catalogue de LM Studio utilise `/v1/models`, celui d’Ollama `/api/tags`.

Tous les modèles d’un catalogue ne prennent pas nécessairement en charge les outils. Choisir un modèle de chat capable d’appeler des fonctions pour agir sur le VBE. Les erreurs de quota, d’authentification ou de modèle sont signalées lors de l’appel. Les offres web Claude/Gemini ne sont pas utilisées comme identifiants API.

## GitHub Copilot

Installer le CLI GitHub Copilot et utiliser **Se connecter à GitHub** dans les paramètres, en utilisant son espace privé. Le complément utilise l’authentification gérée par ce CLI ; il n’extrait pas de jeton d’une extension Visual Studio et ne remplace pas Copilot par GitHub Models.

Le complément lance `copilot.exe --headless --stdio --no-auto-update --log-level error`. Si l’exécutable natif n’est pas dans PATH, renseigner son chemin dans `CODEXVBE_COPILOT_CLI`. Les lanceurs `.cmd` ne sont pas exécutés par un shell. Le catalogue provient de `models.list` ; le transport vérifie les versions de protocole 2 ou 3 et refuse les autres avec un diagnostic explicite.

Une session Copilot est créée pour chaque envoi avec l’historique local et les instructions système. Les appels d’outils et leurs résultats rejoignent cet historique. Aucun identifiant de session distante Copilot n’est réutilisé entre documents. Le CLI conserve son état dans le dossier privé Copilot de VBAi, distinct du dossier des autres clients et du stockage SQLite du complément.

Seuls les outils VBA déclarés sont exposés. Les demandes natives shell/fichiers/réseau/MCP sont refusées. Une permission `custom-tool` correspondant à un outil enregistré peut seulement acheminer l’appel vers `LlmVbeTools.InvokeAsync`, qui conserve les contrôles Discussion/Plan, Lecture seule, portée du projet, approbation configurée et révision SHA. Les appels dupliqués ne sont pas exécutés deux fois. Arrêter ferme le processus appartenant à ce client ; une action COM déjà commencée peut se terminer et conserve son rollback.

## Vérification

```powershell
dotnet build tests/CodexVBE.Providers.Smoke/CodexVBE.Providers.Smoke.csproj -c Debug -p:BuildOutputRoot="$PWD/artifacts/provider-tests"
./artifacts/provider-tests/CodexVBE.Providers.Smoke/Debug/net48/ProviderTests.exe
```

Les tests utilisent des gestionnaires HTTP simulés et un processus CLI de test avec le véritable cadrage `Content-Length` : URLs, en-têtes, catalogues, pagination, appels/résultats d’outils, Unicode, métadonnées de continuation, stockage DPAPI, migration des réglages, brouillons du formulaire et annulation. Copilot couvre les protocoles 2/3, les permissions natives refusées, les outils enregistrés, les doublons, les sessions étrangères, une version inconnue et l’arrêt d’une réponse.

Les fournisseurs compatibles et Claude utilisent le streaming SSE lorsqu’ils sont appelés depuis le chat ; Copilot utilise ses notifications de texte. Les arguments d’outils fragmentés sont assemblés avant exécution. Une fin de flux absente, une réponse tronquée ou une erreur ne déclenche pas d’appel partiel. Les signatures et métadonnées de continuation restent dans l’historique protocolaire ; le raisonnement brut n’est pas affiché. Une réponse JSON complète reste acceptée lorsqu’un serveur ignore le streaming. Bedrock affiche actuellement la réponse complète de Converse, sans ConverseStream. Les résumés de réflexion progressifs restent propres à Codex.

**Validation réelle OpenRouter, 27 septembre 2026 : PASS.** Avec la clé présente dans l’environnement et le routeur `openrouter/free`, le catalogue a retourné 458 modèles ; une réponse synthétique a été reçue en streaming. Un second essai a reçu un appel d’outil fictif, renvoyé le nombre 42 et obtenu la réponse finale correspondante. Aucun code du projet ni contenu VBA n’a été transmis, aucune action VBE n’a été exécutée. Commandes explicites : `ProviderTests.exe --live-openrouter` et `ProviderTests.exe --live-openrouter-tools` ; elles ne font pas partie des tests locaux par défaut.

**Autres fournisseurs : validation authentifiée NOT_RUN.** Les tests locaux couvrent leurs protocoles, les modes clé/Entra Azure, les déploiements, les modèles manuels, Converse, SSE, annulation et contenus signés. Le CLI Copilot réel n’était pas disponible dans PATH. Les essais ne constituent pas une validation dans un hôte VBE réel.

## Contrats utilisés

- [Anthropic : appels et résultats d’outils](https://platform.claude.com/docs/en/agents-and-tools/tool-use/handle-tool-calls), [catalogue](https://platform.claude.com/docs/en/api/models/list).
- [Gemini : compatibilité OpenAI](https://ai.google.dev/gemini-api/docs/openai).
- [Mistral : API](https://docs.mistral.ai/api).
- [DeepSeek : catalogue](https://api-docs.deepseek.com/api/list-models/).
- [OpenRouter : Chat Completions](https://openrouter.ai/docs/quickstart).
- [LM Studio : outils compatibles OpenAI](https://lmstudio.ai/docs/developer/openai-compat/tools).
- [GitHub Copilot CLI](https://docs.github.com/en/copilot/reference/copilot-cli-reference/cli-command-reference), [transport du SDK officiel](https://github.com/github/copilot-sdk/blob/main/python/copilot/client.py), [appels et notifications du SDK](https://github.com/github/copilot-sdk/blob/main/python/copilot/session.py).
- [Azure : endpoints et déploiements](https://learn.microsoft.com/en-us/azure/ai-studio/ai-services/concepts/endpoints), [authentification](https://learn.microsoft.com/en-us/azure/api-management/api-management-authenticate-authorize-ai-apis).
- [xAI : API](https://docs.x.ai/developers/rest-api-reference/inference), [Groq : outils en streaming](https://console.groq.com/docs/tool-use/local-tool-calling).
- [Bedrock : Converse](https://docs.aws.amazon.com/bedrock/latest/APIReference/API_runtime_Converse.html), [clés API Bedrock](https://docs.aws.amazon.com/bedrock/latest/userguide/api-keys-use.html).
