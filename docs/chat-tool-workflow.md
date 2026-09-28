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

## Pause après huit réponses avec outils

La boucle HTTP dispose d'un budget de huit réponses du modèle par segment. Si le modèle demande encore des outils lors de la huitième réponse, leurs résultats sont enregistrés puis le tour passe en pause, sans exception de dépassement.

- Le transcript présente les outils terminés/refusés et indique ce qui reste : réponse finale, plus vérification automatique lorsque cette option est activée et que du code a changé.
- Utiliser **Reprendre le tour en pause** dans le menu, ou **Reprendre ▶** avec un champ de saisie vide.
- La reprise conserve le même identifiant de tour, les messages et chaque résultat d'outil. Elle n'ajoute pas une nouvelle question et ne recommence pas le travail précédent.
- Un identifiant d'appel déjà enregistré ne déclenche jamais une deuxième exécution, même si le modèle le renvoie.
- Fournisseur, modèle, effort et mode doivent correspondre au profil enregistré lors de la pause. Le projet et les permissions sont revalidés, et les instructions système actualisées.
- Après la reprise, les nouveaux changements disposent de leur compte rendu et du rollback existant. La vérification automatique attend la fin du tour complet.

L'état de pause, le profil et l'historique sont persistés avec la session. Un nouvel envoi de texte démarre une nouvelle demande et conserve l'historique précédent. En cas d'interruption, un outil sans résultat enregistré est marqué comme non confirmé : inspecter l'état réel avant toute nouvelle action.

Ce budget appartient à la boucle HTTP de VBAi. Codex et les transports SDK gèrent leur orchestration interne ; cette modification n'impose pas artificiellement huit appels à ces moteurs.

## Validation

Les tests de catalogue vérifient la disponibilité du catalogue complet par découverte, le filtrage Discussion et l'absence de contournement des restrictions de projet. Le scénario HTTP simule huit actions, une pause, un profil modifié refusé, puis une reprise comprenant un identifiant déjà traité : huit exécutions seulement, résultats et identifiant de tour conservés.
