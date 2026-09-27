# Couverture du client Codex app-server

Le lot global du 27 septembre 2026 est passé : 219 tests réussis, 2 ignorés. Son rapport est `artifacts/coverage/batch-comprehensive-final/9bf3adb1-a9d7-4eac-bc7c-3913122c6063/coverage.cobertura.xml`. Les tests du client emploient un faux transport ; aucun processus Codex authentifié n'a été lancé pour ces scénarios.

Le transport JSONL est injecté par `ICodexAppServerTransport`. En production, `CodexProcessTransport` garde le lancement `codex app-server`, la sortie/erreur redirigées et l'écriture UTF-8 sans BOM. Les tests utilisent uniquement un faux en mémoire.

| Parcours | Scénarios écrits | Limite restante |
| --- | --- | --- |
| Initialisation et compte | Ordre `initialize`/`initialized`/`account/read`, refus d'un compte non ChatGPT, échec de démarrage, absence d'identité de thread | Localisation de l'exécutable, environnement réel, OAuth ChatGPT et version effective du protocole Codex. |
| Conversation | Nouveau thread avec `sandbox=read-only` et `approvalPolicy=untrusted`, reprise par ID exact, notification `ThreadReady` | Reprise réelle d'un thread sauvegardé et invalidation distante. |
| Catalogue | Pages et curseur, modèles par défaut, efforts disponibles, données manquantes, erreur JSON-RPC | Catalogue courant, pagination longue et changements de schéma réels. |
| Tour et streaming | Démarrage, delta message, résumé et partie de résumé, message final, fallback sans texte, erreur de tour | Flux réel sur une session Codex authentifiée et ordre concurrent des notifications. |
| Annulation et arrêt | `turn/interrupt` avec ID exact, statut interrompu, refus de tours parallèles, arrêt du transport et `Dispose` avec tour en attente | Terminaison/vidage des vrais flux OS lors d'une fermeture non coopérative. |
| Protocole défensif | Requête serveur inconnue `-32601`, notification étrangère, appel d'outil étranger sans invocation VBIDE, JSON malformé | Appel d'outil valide avec résultat, approbation interactive et charge utile très grande ; à exercer avec l'hôte réel et les tests d'outils dédiés. |

Les branches propres à `ProcessStartInfo`, au pipe standard, à la BOM et au signal de sortie du système d'exploitation ne sont pas couvertes par le faux transport. Une couverture complète de ces branches réclame un vrai processus enfant contrôlé et une vérification séparée du binaire Codex installé ; le faux ne représente pas l'authentification ni les réponses du service réel.
