# Confidentialité, sauvegarde et robustesse du chat

## Modifications

| Point | Résultat |
| --- | --- |
| A — Confidentialité | Lectures ciblées contrôlées comme les écritures ; accès supplémentaires en lecture et contexte global autorisés distinctement dans une fenêtre Designer. Inventaires filtrés, emplacement du débogueur attribué au volet réel, références contrôlées. Anciennes sessions migrées vers un contexte fournisseur vierge ; historique local conservé et exclu des branches futures. |
| B — Synchronisation et sauvegarde | Synchronisation automatique inchangée. Ctrl+S synchronise puis exécute Enregistrer dans le VBE sur le projet de l'onglet. Succès vérifié sur le projet et le document hôte ; état inconnu présenté comme non vérifié. |
| C — Catalogue | Noyau de 8 outils ; découverte par familles et passerelle vers les 206 fonctions. HTTP transmet au maximum 64 schémas, prioritaires pour la dernière famille. Discussion/Plan filtrent les mutations. |
| D — Budget | Pause après 8 réponses HTTP avec outils ; bilan, profil et résultats persistés, reprise du même tour, refus de rejouer un identifiant déjà traité et vérification finale des nouveaux changements. |
| E — Pont | Lecture UTF-8 bornée à 10 Mio et 10 secondes avant désérialisation ; récupération après client silencieux, interrompu ou mal formé. |

Les définitions initiales représentent 3 025 caractères JSON contre 127 300 pour le catalogue complet, soit environ 97,6 % de réduction. Ce résultat est une mesure statique de schémas, pas un coût facturé ou un temps d'intervention mesuré chez un fournisseur.

## Validation locale

- Compilation finale : 0 erreur, 0 avertissement.
- Passe élargie : 331 tests réussis, aucun ignoré (`artifacts/review-final/review-final-v2.trx`).
- Recontrôle après la dernière modification de migration/branche : 84 tests réussis, aucun ignoré (`artifacts/review-final/privacy-workflow-final.trx`). Ces tests recoupent la passe précédente ; les nombres ne s'additionnent pas.
- Excel : vrai Ctrl+S, contrôle de l'état du classeur, cible correcte malgré un autre classeur actif, fermeture et réouverture du fichier avec le code attendu. Le routage natif est effectué une seule fois (`artifacts/review-final/excel-review-final.trx`, résultat du test `MonacoSaveExcelTests`).
- Excel : deux projets ouverts, lecture B refusée puis permise après autorisation, écriture B toujours refusée, inventaires et emplacement de débogage filtrés (`artifacts/review-final/excel-privacy-final.trx`, 1 réussite, aucun ignoré). Ce scénario a été relancé séparément après la sortie d'Excel : le premier lancement groupé l'avait ignoré car le processus du scénario de sauvegarde n'avait pas encore terminé.
- JavaScript : 11 tests de langage Monaco réussis.

L'annulation du premier enregistrement, les erreurs et l'état hôte inconnu sont simulés à la frontière native ; le dialogue Enregistrer sous n'est pas certifié par ces simulations. Les validations réelles concernent Excel ; les tests SOLIDWORKS restent différés. Les échanges des fournisseurs sont simulés pour le budget et le catalogue. Il ne s'agit pas d'une mesure de couverture globale du dépôt.

La build locale enregistrée est mise à jour dans `artifacts/monaco/host-build` du worktree, avec sauvegarde de la précédente dans `artifacts/review-final/host-build-before.zip`. Les clés COM gardent leur emplacement existant ; aucun binaire de main n'est remplacé.

## Documentation détaillée

- [Confidentialité des projets](project-privacy.md)
- [Catalogue et reprise](chat-tool-workflow.md)
- [Enregistrement Monaco](modern-editor.md)
- [Budget du pont local](bridge-request-budget.md)

## Évolution du compositeur et de la pause de secours

Le compositeur fonctionne pendant les interventions : arrêt à vide, mise en attente avec du texte, file persistée par session et actions Envoyer maintenant / Modifier / Supprimer. Le départ immédiat attend la fin de l'interruption ; la file ne repart pas toute seule après un arrêt manuel, une erreur ou une pause. Les références et sélections sont revalidées avant leur transmission et les brouillons en cours sont conservés.

Les huit tours représentent désormais un secours après huit réponses consécutives sans nouveau résultat réussi. Une intervention qui progresse continue ; un plafond séparé de 64 réponses protège les boucles dont les résultats changent continuellement. Voir [le workflow](chat-tool-workflow.md) pour la définition précise de la progression et les limites.

Validation : build sans avertissement ni erreur ; 141 tests ciblés chat, sessions, localisation et Designer réussis, zéro ignoré (`artifacts/queue/results/queue-final.trx`). Les scénarios couvrent l'ordre de la file, la conservation du brouillon, l'interruption immédiate, son échec, l'arrêt manuel, la modification/suppression, la sérialisation des messages et de leur contexte, la progression au-delà de huit réponses et le plafond de sécurité. La vue WinForms de la file a été inspectée dans les thèmes clair et sombre, à 460 et 720 pixels ; les captures isolées des contrôles sont dans `artifacts/queue/render/queue-panel-*.png`. Ces validations utilisent les fournisseurs simulés et ne constituent pas un essai réseau avec chaque fournisseur.
