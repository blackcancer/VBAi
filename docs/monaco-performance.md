# Réactivité Monaco : diagnostic et plan d’optimisation

## Périmètre de l’étude initiale

Étude du commit `c67ffc9`, le 29 septembre 2026. Lecture du code, compilation Debug et mesure isolée sur des sources synthétiques. Aucun changement du comportement de production, aucune exécution de macro et aucune qualification native Excel/SOLIDWORKS dans cette étude.

## Constats établis dans le code

### Débogage et points d’arrêt

- `ModernEditorWindow.Designer.cs:119` fixe le timer à **900 ms**. `TimerTick` effectue la synchronisation de tous les documents avant `ObserveDebugMode` (`ModernEditorWindow.cs:345`).
- Après une commande de pas à pas, `EditorCommand` efface les marqueurs et invalide le mode mémorisé ; la nouvelle ligne attend le passage du timer. Le délai de planification peut approcher 900 ms, auquel s’ajoutent le travail de synchronisation et la disponibilité du thread hôte. Ce n’est pas une borne maximale de latence réelle.
- `EditorCommand` capture tous les documents, puis appelle `ProcessDocumentsCore(true)`, qui les capture une seconde fois (`ModernEditorWindow.Debug.cs:67-71`, `ModernEditorWindow.cs:314`). Cette deuxième capture est actuellement une barrière de fraîcheur : sa suppression exige de conserver la détection des modifications survenues pendant les attentes asynchrones.
- La découverte d’une commande parcourt jusqu’à dix pages de 200 entrées. Chaque `ListCommands` reconstruit l’inventaire complet des menus, avant de le paginer. `InvokeCommand` le parcourt encore. Le chemin effectue donc de deux à onze inventaires si la commande est trouvée, selon sa position (`VbeDebug.cs:65,412,797`).
- Le marqueur de point d’arrêt est envoyé seulement après l’invocation native. Il représente une demande non vérifiée, car VBIDE n’expose pas l’inventaire des points d’arrêt. Un retour visuel rapide doit conserver cette distinction.
- Le nettoyage des lignes d’exécution utilise un appel WebView séquentiel par document. Un message groupé pourrait porter la destination et les marqueurs à retirer.

### Autocomplétion, survol et signatures

- `language.js:11-25` mutualise seulement les requêtes simultanées de même révision. Dès la réponse reçue, une nouvelle demande repart vers l’hôte même si le code n’a pas changé. Cette décision assure actuellement la fraîcheur des références.
- Chaque `LanguageRequest` capture les onglets, relit tous les composants du projet via COM, remplace les sources ouvertes par leurs brouillons, reconstruit l’index du projet et retourne **tous les symboles et toutes les sources** (`ModernEditorWindow.Language.cs:18-44`).
- Les références possèdent déjà un cache, limité à la dernière clé sur le thread courant. Un défaut de cache peut effectuer jusqu’à six passages sur les bibliothèques pour développer les types dépendants (`EditorReferenceIndex.cs:31-93`).
- `provideCompletionItems` prépare la documentation de chaque suggestion. L’API Monaco embarquée permet de la compléter à la demande avec `resolveCompletionItem`; elle fournit aussi un `CancellationToken`, non consommé par les fournisseurs actuels.
- Le langage et les plans de synchronisation utilisent le même worker FIFO. Une indexation de bibliothèque longue peut retarder un plan requis avant une commande native (`EditorSyncWorker.cs:50-74`).

### Synchronisation et communication

- Chaque changement de modèle transmet le texte complet du module (`editor.js:72`). Les captures ultérieures retransmettent tous les textes ouverts.
- Le cycle périodique prépare un plan pour chaque document, même propre. Pour un document propre, le magasin ne chiffre ni ne réécrit de brouillon ; le passage dans la file et le calcul/contrôle du texte subsistent.
- La synchronisation attend 600 ms de repos mais est échantillonnée par le timer de 900 ms : sans charge supplémentaire, le départ peut se situer environ 600 à 1 500 ms après la dernière frappe. Ce délai concerne la synchronisation automatique, pas le déclenchement des complétions.
- `Script` sérialise ses arguments sur le thread UI avant `ExecuteScriptAsync`. Le coût du gros objet de langage s’ajoute donc au travail du thread qui sert les interactions.

## Mesure reproductible

Script : `tools/probes/Measure-MonacoIndex.ps1`.
Résultat local : `artifacts/monaco-performance/index-baseline.json` (date, commit et SHA-256 de l’assemblage inclus).

Chaque module contient 101 lignes : `Option Explicit` et 25 procédures avec paramètre et variable locale. Trois passages de chauffe puis 30 mesures par taille. L’index est celui de production ; l’enveloppe JSON contient les mêmes catégories de données que la réponse de langage. L’indexation et la sérialisation sont mesurées séparément avec PowerShell/.NET Framework ; ces médianes ne constituent pas une mesure de latence de bout en bout.

| Modules | Lignes | Symboles | Index p50 / p95 | JSON p50 / p95 | Réponse UTF-8 |
| --- | ---: | ---: | ---: | ---: | ---: |
| 1 | 101 | 76 | 0,75 / 0,90 ms | 2,52 / 3,27 ms | 33 045 octets |
| 10 | 1 010 | 760 | 6,26 / 8,53 ms | 23,40 / 31,80 ms | 329 901 octets |
| 30 | 3 030 | 2 280 | 25,85 / 51,94 ms | 84,86 / 127,47 ms | 991 161 octets |
| 100 | 10 100 | 7 600 | 99,55 / 135,80 ms | 350,54 / 426,37 ms | 3 305 571 octets |

Ces chiffres excluent les lectures COM, les bibliothèques externes, les attentes de file, le transport WebView2 et le rendu des suggestions. Les 100 modules représentent un projet indexé, pas 100 onglets ouverts (limite actuelle : 30). Aucun gain après optimisation n’est encore mesuré.

## Ordre de mise en œuvre recommandé

### 1. Réponse aux commandes de débogage

- Instrumenter une requête depuis le clic/F9/F8 jusqu’au rendu : capture, attente de file, synchronisation, recherche native, exécution, observation, rendu. Conserver seulement durées, compteurs et identifiants éphémères dans les traces, sans code VBA.
- Déclencher une observation après la commande, hors callback WebView, puis des vérifications légères bornées si le VBE traite la commande de manière différée. Garder une surveillance moins fréquente pour les actions déclenchées directement dans le VBE.
- Séparer cette observation du traitement périodique de tous les brouillons. Regrouper les mises à jour des marqueurs dans un seul message.
- Rechercher les commandes natives par identifiant lorsqu’il est connu, en validant à chaque invocation le mode, l’activation, le libellé autorisé et `Enabled`. Garder un repli vers l’inventaire ; ne pas mémoriser durablement un objet COM ou son état activé.
- Utiliser une file ordonnée pour les commandes. Ne jamais fusionner arbitrairement deux toggles de point d’arrêt ou deux pas à pas : leur multiplicité a un sens.
- Afficher immédiatement un état « demande en cours », puis le retirer en cas de refus ; ne pas transformer cet état en preuve d’un point d’arrêt confirmé.

### 2. Index de langage incrémental

- Cache mémoire isolé par identité de projet, composant et révision ; reconstruire seulement les modules modifiés, puis composer un instantané immuable du projet.
- Invalider sur édition Monaco, édition par outils/Git, import, renommage, retrait de composant et changement de références. Réutiliser les observateurs `VbeCollectionEvents` et `VbeReferenceEvents`, avec réconciliation bornée des modifications natives non couvertes. Les abonnements peuvent échouer selon l’hôte : prévoir un repli explicite.
- Précharger les métadonnées des références hors du chemin d’une frappe et conserver un cache borné par bibliothèque/version/types. Dédier le travail de langage à une file distincte du travail de récupération/synchronisation, sans transmettre de proxies VBE au worker.
- Inclure une génération de projet et une révision de document dans le protocole. Annuler ou abandonner les demandes obsolètes, y compris celles d’un document fermé ou d’un ancien projet.
- Envoyer un index initial puis ses changements ; récupérer les sources de définition à la demande. Charger la documentation détaillée uniquement pour la suggestion sélectionnée.
- SQLite n’est pas nécessaire sur ce chemin interactif. Un éventuel cache persistant de métadonnées pourra être évalué séparément pour le démarrage à froid.

### 3. Synchronisation et messages plus petits

- Ne préparer un plan de modification que pour un document modifié, tout en continuant à observer les changements natifs des documents propres.
- Remplacer les snapshots systématiques par des changements versionnés, avec resynchronisation complète en cas de trou de révision. Conserver une barrière explicite avant sauvegarde, fermeture et commandes natives.
- Avant une commande native, synchroniser les brouillons nécessaires du projet et revalider mode, révision et empreinte juste avant l’action. Un document propre ne dispense pas de vérifier la source native.
- Éviter les réécritures répétées du même brouillon, sans affaiblir sa récupération après erreur. Réserver le passage sur disque à un nouvel état à préserver.
- Évaluer des notifications groupées via `PostWebMessageAsJson` pour les mises à jour sans résultat. Ce changement ne supprime pas à lui seul le coût de sérialisation : la réduction des données est prioritaire.

## Contraintes et critères de validation

WebView2 reste sur son thread UI STA ; les accès VBE restent sur leur thread propriétaire. Les calculs sur instantanés immuables peuvent sortir de ce thread. Les sorties de callbacks via `Task.Yield`/publication UI protègent de la réentrance et ne doivent pas être supprimées aveuglément.

Cibles initiales proposées, non garanties : retour visuel de commande sous 50 ms ; complétion à chaud p95 sous 100 ms ; rafraîchissement de ligne sous 100 ms après disponibilité du nouvel état natif. Mesurer séparément les appels natifs lents, le démarrage à froid et le temps d’exécution de la macro, qui ne sont pas des coûts de peinture Monaco.

Comparer p50/p95, octets échangés, lectures COM, reconstructions d’index et blocages UI sur projets de 1/10/30/100 modules, avec 1/10/30 onglets. Couvrir frappe rapide, F8 répété, bascule de point d’arrêt, erreur de compilation, correction en pause, références ajoutées/retirées, conflit natif, changement de projet, fermeture et reprise d’un brouillon. Excel et SOLIDWORKS demandent des qualifications séparées ; les micro-mesures ci-dessus ne les remplacent pas.

Les éventuels timers et contrôles fixes restent déclarés dans le Designer WinForms.

## Sources techniques

- [Microsoft — performances WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/performance) : réduire les échanges, grouper les messages, profiler les scénarios réels.
- [Microsoft — modèle de threads WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/threading-model) : affinité STA/UI, réentrance et traitements asynchrones.
- Contrat de la version Monaco embarquée : `artifacts/monaco-dependencies/monaco/package/monaco.d.ts:7449-7461`, `CompletionItemProvider`, paramètre d’annulation et résolution différée des détails.


## Implémentation et qualification

Les modifications suivantes sont maintenant réalisées dans le worktree `chat-ux` :

- **Flux de saisie** : chaque frappe envoie ses changements UTF-16 avec la révision de base et la nouvelle révision. Le récepteur valide les plages avant application atomique. Une révision manquante ou une plage invalide déclenche une capture complète du modèle courant ; une réponse ancienne ne remplace jamais une révision plus récente.
- **Écriture continue** : minuterie Designer de 40 ms, regroupement après 120 ms de repos ou 450 ms de saisie soutenue. Ces valeurs sont des seuils de planification et non des garanties de délai réel. Seuls les documents modifiés sont préparés dans ce parcours. Le cycle périodique reste disponible pour observer le code natif et reprendre les modifications différées.
- **Récupération et conflits** : le brouillon est conservé avant les écritures ; le plan et la source native sont revalidés. Les restrictions natives en exécution/pause, l’encodage et les attributs VBA restent contrôlés. Une écriture refusée attend une nouvelle frappe ou le cycle de réconciliation ; elle ne détruit pas le brouillon.
- **Captures différentielles** : `snapshots` reçoit les révisions connues et ne renvoie que les modèles différents. Les captures explicites avant sauvegarde/fermeture restent présentes.
- **Débogage** : observation indépendante toutes les 125 ms et observation publiée immédiatement après un pas. Une file conserve les commandes successives. La capture redondante a été retirée avec revalidation avant exécution. Les commandes connues utilisent `FindControl` avec vérification de leur identité, libellé et état ; le repli fait un inventaire borné, sans pagination répétée.
- **Retour visuel** : une demande de point d’arrêt affiche immédiatement un indicateur temporaire, retiré après traitement. Le marqueur final reste une demande non vérifiée. Les lignes d’exécution sont mises à jour en un seul appel WebView.
- **Langage** : worker séparé, cache borné des déclarations par contenu/type/nom de module, réponses divisées en blocs indépendamment versionnés pour modules et références. Un projet inchangé renvoie un acquittement compact ; un module changé ne retransmet pas les autres blocs. Les sources ne sont transmises que pour les définitions. La documentation des suggestions est résolue à la demande et les demandes obsolètes sont abandonnées.
- **Fraîcheur** : les sources natives et la liste des références restent relues avant validation de l’index. Le cache ne repose donc pas sur des événements COM qui pourraient manquer, ni sur un délai arbitraire d’expiration. Le cache de bibliothèques existant reste sur le worker langage. Le coût résiduel des lectures COM et du premier chargement des bibliothèques doit être mesuré dans chaque hôte.
- **Mesures locales** : `PerformanceSample` expose les durées de sérialisation, appels WebView et commandes, ainsi que la taille des scripts en caractères, sans transmettre leur contenu. Aucun envoi de télémétrie n’est ajouté.
- **WinForms** : les deux nouvelles minuteries sont dans `ModernEditorWindow.Designer.cs`, leurs gestionnaires dans le fichier principal. Le démarrage est publié via le dispatcher WinForms après les callbacks WebView2. Le test réel a détecté et permis de corriger un démarrage qui affichait une minuterie active sans vidanger le flux.

### Comparaison synthétique après modification

Même corpus de 30 modules, mêmes 30 itérations après chauffe. Les valeurs ci-dessous concernent l’index géré et l’enveloppe JSON, pas le trajet natif complet. Dans le scénario « un module édité », une ligne de commentaire différente est ajoutée à chaque itération.

| Scénario | Index p50 | Sérialisation p50 | Réponse |
| --- | ---: | ---: | ---: |
| Avant : reconstruction/réponse complète | 25,85 ms | 84,86 ms | 991 161 octets |
| Après : projet inchangé | 0,43 ms | 0,01 ms | 101 octets |
| Après : un module édité | 1,41 ms | 2,10 ms | 33 910 octets |

Preuves locales : `artifacts/monaco-performance/index-baseline.json`, `index-unchanged.json` et `index-edit-one.json`. La taille exacte en production varie notamment avec les identifiants, les déclarations et les bibliothèques ; aucune accélération globale chiffrée dans Excel n’est déduite de ces micro-mesures.

### Validations

- 30 tests JavaScript : résolution des objets, références ajoutées/retirées, blocs incrémentaux, annulation et documentation différée.
- 176 tests .NET de régression éditeur/débogage/langage et WebView2 réussis (`accepted.trx`), complétés par trois tests des chemins rapides (`fast-paths.trx`). Ces trois tests sont distincts du lot de 176.
- Le vrai moteur WebView2/Monaco transmet les deltas, synchronise automatiquement un adaptateur de module jetable, récupère une révision volontairement perdue et applique les marqueurs groupés. Ce scénario ne constitue pas une preuve COM Excel.
- 46 surfaces Designer chargées, redimensionnées et sérialisées (`artifacts/monaco-performance/designers`).
- Aucun test de macro dans l’instance Excel existante, conservée intacte ; pas de qualification SOLIDWORKS pour ce lot. Couverture globale non recalculée.
