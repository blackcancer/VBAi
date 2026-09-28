# Catalogue progressif et reprise des tours

## Outils par famille

Le modèle commence avec un noyau de huit outils : état, projets, modules, lecture native, ouverture/lecture Monaco et les deux passerelles de catalogue. Le catalogue complet reste disponible avec `discover_tools` :

| Famille | Contenu |
| --- | --- |
| `code` | Code, procédures, Monaco, navigation et refactorings |
| `forms` | UserForms, contrôles et concepteur |
| `debug` | Compilation, débogueur et commandes d'exécution |
| `git` | Git, branches, checkpoints et GitHub |
| `environment` | Projets, références, options et environnement VBE |

`Family=all` retourne tous les schémas autorisés par le mode. Discussion et Plan ne présentent que les outils d'inspection ; Agent expose également les mutations. La découverte décrit les capacités, sans accorder de permission supplémentaire sur un projet.

Pour HTTP, une famille découverte est ajoutée aux définitions transmises aux réponses suivantes, avec un plafond de 64 schémas et priorité au noyau. Les autres fonctions restent accessibles par la passerelle, y compris après Family=all. Pour Codex, les outils dynamiques étant enregistrés au démarrage du fil, le noyau reste stable et `invoke_tool(ToolName, ArgumentsJson)` appelle une fonction découverte. Cette passerelle repasse systématiquement par les gardes normales : mode, projet lié, accès aux autres projets, contexte partagé et politique d'édition. Les appels récursifs de passerelle sont refusés.

## Pause de secours

La boucle HTTP continue au-delà de huit réponses lorsque les outils apportent de nouveaux résultats réussis. Huit réponses consécutives avec outils sans progression déclenchent une pause de secours : refus, erreurs, identifiants déjà traités ou résultats réussis identiques. La progression est évaluée sur le triplet nom/arguments/résultat ; ce n'est pas une mesure de l'accomplissement du besoin métier. Un plafond distinct de 64 réponses par segment borne également les boucles dont les résultats changent constamment. Les résultats sont enregistrés avant la pause, sans exception de dépassement.

- Le transcript présente les outils terminés/refusés et indique ce qui reste : réponse finale, plus vérification automatique lorsque cette option est activée et que du code a changé.
- Utiliser **Reprendre le tour en pause** dans le menu, ou **Reprendre ▶** avec un champ de saisie vide.
- La reprise conserve le même identifiant de tour, les messages et chaque résultat d'outil. Elle n'ajoute pas une nouvelle question et ne recommence pas le travail précédent.
- Un identifiant d'appel déjà enregistré ne déclenche jamais une deuxième exécution, même si le modèle le renvoie.
- Fournisseur, modèle, effort et mode doivent correspondre au profil enregistré lors de la pause. Le projet et les permissions sont revalidés, et les instructions système actualisées.
- Après la reprise, les nouveaux changements disposent de leur compte rendu et du rollback existant. La vérification automatique attend la fin du tour complet.

L'état de pause, le profil et l'historique sont persistés avec la session. Un nouvel envoi de texte démarre une nouvelle demande et conserve l'historique précédent. En cas d'interruption, un outil sans résultat enregistré est marqué comme non confirmé : inspecter l'état réel avant toute nouvelle action.

Ce budget appartient à la boucle HTTP de VBAi. Codex et les transports SDK gèrent leur orchestration interne ; cette modification n'impose pas artificiellement huit appels à ces moteurs.

## Compositeur pendant une intervention

- Champ vide : le bouton devient **Arrêter ■**. Entrée demande également l'arrêt ; Maj+Entrée conserve le retour à la ligne.
- Champ rempli : **Mettre en attente ↑** (ou Entrée) capture le texte, les références, sélections et mémoire ; la réponse actuelle continue.
- La file s'affiche immédiatement au-dessus du champ. Chaque message dispose de **Envoyer maintenant**, **Modifier** et **Supprimer**, avec info-bulles et thème.
- Les messages sont traités dans l'ordre après une réponse terminée. Leur projet, permissions et révisions de code sont revérifiés au départ ; un contexte périmé bloque le départ sans retirer le message.
- **Envoyer maintenant** donne priorité au message choisi, demande l'interruption et attend le nettoyage du tour avant de démarrer la nouvelle demande. Les autres messages suivent après son succès. Si l'interruption échoue, il reste en attente.
- **Modifier** remet le message et son contexte dans le compositeur. Un brouillon déjà présent est préservé ; il faut l'envoyer ou le vider avant cette action. Le message modifié peut ensuite être remis en attente.
- Un arrêt manuel, une erreur ou une pause de secours laisse les autres messages en attente. Le brouillon en cours n'est pas remplacé par l'envoi automatique.
- La file appartient à la session liée au projet et est persistée dans SQLite. La réouverture l'affiche sans l'exécuter automatiquement.

## Validation

Les tests de catalogue vérifient la disponibilité du catalogue complet par découverte, le filtrage Discussion et l'absence de contournement des restrictions de projet. Le scénario HTTP simule huit actions sans progression, une pause, un profil modifié refusé, puis une reprise comprenant un identifiant déjà traité : huit exécutions seulement, résultats et identifiant de tour conservés.
