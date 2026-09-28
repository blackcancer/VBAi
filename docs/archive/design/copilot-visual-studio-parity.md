# Assistant CodexVBE : cible inspirée de Copilot dans Visual Studio

> Archive conservée le 28 septembre 2026. Ce document contient des observations et des décisions de sa période de rédaction ; ses états « à faire » et ses anciens chiffres ne constituent pas le bilan actuel. Voir [la documentation actuelle](../../README.md) et [les travaux restants](../../roadmap.md).

État de la proposition : 27 septembre 2026. Les fonctionnalités de Copilot varient selon la version de Visual Studio et le compte ; cette liste décrit la cible produit de CodexVBE, pas une compatibilité déjà acquise. Le VBE reste la source de vérité pour le code VBA vivant.

Tranche implémentée localement : saisie Entrée pour envoyer et Maj+Entrée pour une nouvelle ligne, correction orthographique française WPF dans le champ, et suggestions `#` pour projet et module, et `@` pour Sub, Function et Property Get/Let/Set. La sélection ajoute le code VBA vivant au message lors de l'envoi ; une procédure dont le SHA a changé doit être sélectionnée à nouveau. L'index se construit par étapes sur le thread VBE. La conversation affiche des cartes et des blocs de code copiables. `replace_lines` applique le code sans dialogue préalable, puis affiche une carte avec le diff et une annulation contrôlée par SHA. La compilation et les essais isolés passent ; dans Excel, l'ouverture du chat, le catalogue Codex et les sélections initiales ont été vérifiés. Les autres parcours de ce tableau gardent leur propre état de validation.

## Expérience cible

La tranche actuelle est détaillée dans [Interface de conversation et sessions VBA](../../chat-ui.md) : surface WPF, diff et résumés inline, références `#` et `@`, historique multichat par document et mémoire locale SQLite.

| Domaine | Expérience attendue dans CodexVBE | État du dépôt |
| --- | --- | --- |
| Conversation | Panneau ancrable, réponses progressivement affichées, blocs de code copiables, navigation vers les cibles, états des outils, arrêt d'une réponse et relance | Surface WPF et panneau natif à droite testés dans Excel ; premier placement droit manuel puis restauré au redémarrage ; streaming Codex, résumés et outils dépliables, arrêt et navigation |
| Modes | Demander (lecture seule), Plan (étapes révisables), Agent (outils et modifications selon la politique configurée) | Modes Discussion / Plan / Agent visibles et persistés ; blocage des mutations et exécutions en Discussion/Plan |
| Références `#` | Liste au curseur dès `#` ; filtrage en cours de frappe ; flèches, Entrée, Tab, Échap ; projets, modules, procédures, Property Get/Let/Set, formulaires et contrôles ; désambiguïsation par projet/module/type | Liste au curseur : # pour projets/modules, @ pour procédures/propriétés ; liens VBE et puces de contexte ; contrôles à venir |
| Contexte | Sélection et module actif implicites et visibles ; ajout explicite de symboles, module, projet, état du débogueur, changements et fichiers ; liste des sources effectivement transmises et budget de contexte | Références, sélection native explicite, aperçu et taille en caractères, détection de sélection périmée ; snapshots transmis conservés |
| Recherche | Recherche par nom, contenu, définition et usages à travers les projets ouverts ; navigation vers la ligne ; classement et actualisation lorsque le VBE change | `find_code`, index progressif en mémoire et navigation par symbole ; pas de recherche d'usages sémantique |
| Édition | Édition immédiate avec diff après application, annulation ciblée et relecture du résultat ; commandes `replace_lines` gardées derrière les contrôles SHA et mode Conception | `replace_lines` sans dialogue, contrôle SHA et lecture après édition ; diffs par blocs, annulation ciblée ou par intervention, prélecture des conflits et conservation des changements indépendants |
| Historique | Conversations multiples, titres, reprise, recherche, contexte et actions conservés ; points de retour des modifications de code avec détection de conflit | SQLite par document, recherche dans les messages/code, épinglage, renommage, archivage, branches de conversation, brouillon et snapshots |
| Commandes | Suggestions `/explain`, `/fix`, `/doc`, `/tests`, `/optimize`, `/generate`, `/review`, `/help` et invites personnalisées ; portée affichée avant envoi | Suggestions /expliquer, /corriger, /refactoriser, /documenter, /tests, /plan avec paramètres libres |
| Débogage | Ajouter au chat sélection, état, pile, Locals, Watches, Immediate et erreurs lisibles ; expliquer puis proposer une correction avec diff | Outils de débogage VBE partiels déjà présents ; pas d'attachement UI unifié |
| Revue | Revue du code sélectionné et des changements de la session, commentaires reliés aux lignes et vérification des corrections | Diffs inline avec navigation et rollback ; commentaires reliés aux lignes à venir |
| Tests et diagnostics | Génération de tests VBA adaptés à l'hôte, analyse de compilation, erreurs et sorties disponibles ; exécution uniquement par des commandes hôte explicitement validées | Compilation manuelle ou après intervention, diagnostic inline, lien vers la sélection en erreur et préparation de correction ; validation native NOT_RUN |
| Fournisseurs | Sélecteurs de fournisseur, modèle et effort ; capacités affichées par modèle, erreurs explicites | Sélecteurs et catalogue Codex testés dans Excel ; choix par défaut et par conversation persistés ; autres fournisseurs partiellement implémentés |
| Personnalisation | Instructions projet et utilisateur, bibliothèque de prompts, agents spécialisés, catalogue d'outils et permissions visibles ; connecteurs MCP comme extension distincte | Notes de mémoire par document dans SQLite, jointes explicitement ; outils typés et politique de modification |
| Médias et externes | Images et fichiers joints avec aperçu, contrôle de transmission, URL explicites si le fournisseur sait les lire | Lecture limitée des fichiers externes sous confirmation ; pas d'image dans le chat |
| Git et partage | Changements et commits attachables au chat quand un dépôt existe ; résumé et revue avant commit ; export d'une conversation | Export Markdown ; intégration Git à venir |
| Intégrations hôte | Actions contextuelles depuis le code VBE et les fenêtres de débogage ; éventuellement chat inline, complétion et suggestion de prochaine édition dans le volet de code | Menus contextuels Expliquer/Corriger/Refactoriser, capture de sélection et hébergement COM du chat dans Excel ; complétion à venir |

Les références de bibliothèque COM demandent un parcours distinct : `list_reference_types` et `list_type_members` sont paginés et peuvent être coûteux. La liste `#` doit d'abord couvrir les symboles du projet VBA ; les types et membres externes viendront via une recherche explicite dans une bibliothèque choisie.

## Contrat de la référence `#`

1. Une suggestion porte une identité structurée : type de cible, projet, module ou formulaire, nom, `ProcKind` si nécessaire, plage de lignes, empreinte du module ou version de l'arbre. Le libellé visible ne suffit jamais à identifier une cible.
2. La liste s'ouvre sous le curseur, suit les déplacements et la mise à l'échelle, filtre sans bloquer le thread VBE et garde le texte saisi si l'utilisateur ferme la liste.
3. La sélection insère un jeton lisible comme `#Projet.Module.Procédure` ; l'envoi résout le jeton structuré et relit la cible vivante. Si le code a changé, la fenêtre l'indique et propose de rafraîchir la référence.
4. Plusieurs occurrences d'un même nom restent distinctes. Property Get, Let et Set sont proposées séparément. Les projets verrouillés ou illisibles apparaissent avec une raison et ne produisent pas de faux contenu.
5. Un `#` tapé dans du texte ordinaire reste du texte si aucune suggestion n'est retenue. Les références sélectionnées sont listées près du champ et sous la réponse, avec la portée réellement transmise au modèle.

## Diff et sécurité des modifications

Le modèle produit une édition ciblée. CodexVBE lit le module actuel, vérifie le SHA de base et la plage, applique `replace_lines`, puis relit le résultat. Une carte dans la conversation donne accès au diff et à l'annulation ; l'historique conserve aussi la modification. Une divergence bloque l'application et demande un nouveau calcul ; elle ne déclenche pas de réessai automatique. Le mode Lecture seule interdit toujours l'écriture. Le mode Demander reste applicable aux autres actions VBE, tandis que `replace_lines` est immédiat pour éviter les validations répétées.

Un point de retour contient le texte exact avant/après, les empreintes, l'identité de la cible et l'identifiant de l'action. Restaurer est une nouvelle écriture VBE : vérifier le SHA courant, prévisualiser le diff inverse et respecter la politique d'approbation. Pour les UserForms, le code et l'arbre de contrôles ont des versions distinctes ; une restauration de conception exige un contrat spécifique avant d'être proposée.

## Index et stockage

Commencer par un index **en mémoire**, constitué à partir des projets ouverts, `VBComponents` et `CodeModule`. L'actualiser par empreinte de module et rafraîchissement explicite ou différé ; ne pas interroger toutes les bibliothèques COM à chaque caractère. Effectuer la lecture COM sur le thread VBE, puis filtrer les données copiées hors COM. L'index doit être borné et ne doit pas retarder la saisie.

SQLite conserve désormais les conversations par document, les points de retour et les notes locales ; la recherche plein texte sur le code reste à venir. Elle n'est pas nécessaire pour afficher la première liste `#`. Avant de l'ajouter, définir emplacement utilisateur, migrations, chiffrement éventuel, durée de conservation et exclusion des secrets. Un cache persistant n'est jamais l'autorité pour appliquer une édition : relire systématiquement le VBE et vérifier la révision.

## Ordre de réalisation proposé

1. **Fondation de l'interface** : composer enrichi, popup accessible au curseur, index en mémoire, références structurées et transmission de sources vérifiables. Valider clavier, souris, DPI, projets homonymes ou verrouillés, modifications entre sélection et envoi.
2. **Revue des éditions** : modèle de proposition, diff par bloc, validation et application avec SHA, liste des modifications, annulation ciblée et points de retour pour le code VBA. Vérifier refus et conflits avant toute mutation.
3. **Conversation complète** : modes Demander/Plan/Agent, streaming, arrêt, sessions et historique, commandes `/`, sources et budget de contexte. Ajouter la persistance locale seulement à cette étape si les besoins de reprise la justifient.
4. **Contexte VBE étendu** : sélection active, formulaires et contrôles, débogage, recherche des usages, actions contextuelles dans l'éditeur et revue liée aux lignes.
5. **Fonctions dépendantes de l'hôte** : complétion inline, suggestion de prochaine édition, capture d'images, Git/commits et intégrations externes, chacune selon les capacités réelles du VBE et de l'hôte. Ne pas annoncer la parité avant leur validation en processus réel.

## Sources de comparaison

- [Contexte et références dans Copilot Chat](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-chat-context-references?view=visualstudio)
- [Agent, outils, diffs et points de retour](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-agent-mode?view=vs-2022)
- [Commandes, instructions et actions contextuelles](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-chat-context?view=visualstudio)
- [Chat et édition inline](https://learn.microsoft.com/en-us/visualstudio/ide/visual-studio-github-copilot-chat?view=visualstudio)
- [Aide au débogage](https://learn.microsoft.com/en-us/visualstudio/debugger/debug-with-copilot?view=visualstudio)
- [Suggestions de prochaine édition](https://learn.microsoft.com/en-us/visualstudio/ide/copilot-next-edit-suggestions?view=visualstudio)
