# Revue du code et GitHub

La fenêtre GitHub conserve le dépôt dans le cache privé du document. Aucun dossier
de travail n’est ajouté à côté du classeur ou de la macro SOLIDWORKS.

## Modifications et historique

Dans **Modifications Git**, cocher les modules à inclure dans le prochain commit.
Un formulaire et son fichier `.frx` constituent une seule unité. Le manifeste est
reconstruit automatiquement ; les références VBA disposent d’une sélection distincte.
Les changements non sélectionnés restent visibles et non commités. Push publie
les commits locaux de la branche configurée.

Le diff propose une vue unifiée ou côte à côte, une recherche, la navigation entre
changements et le repli du contexte. Double-cliquer sur une ligne « … » développe
le contexte. Le rendu virtualisé n’a plus la limite d’affichage de 2 000 lignes.
Pour les très grands modules, le calcul conserve une approximation par bloc afin
de borner sa consommation mémoire ; la restauration reste basée sur le texte exact.

Dans **Historique**, sélectionner un commit pour ses détails ou deux commits pour
comparer leurs sources, puis **Comparer les révisions**. La révision la plus récente
dans la liste constitue la cible. Les 200 commits les plus récents sont affichés.
Les checkpoints peuvent également être consultés. **Restaurer le module sélectionné**
importe uniquement ce module, avec ses ressources, après un checkpoint automatique.
Les autres modifications locales sont conservées. **Comparer** revient au VBA actuel.

**Prévisualiser l’import** récupère la branche distante et affiche les fichiers et
références concernés sans importer. Pull et les autres imports affichent également
ce récapitulatif, créent leur sauvegarde et conservent le rollback existant.
Les références incompatibles ou les modules hôtes manquants empêchent toujours l’import.
La résolution des conflits affiche l’ancêtre commun, les deux versions et le texte
résultant éditable. Les limites existantes des aperçus de conflits (64 Kio par fichier)
restent applicables ; un binaire doit être résolu en choisissant une version complète.

## Dépôts et pull requests

Le nouvel onglet **GitHub** utilise le compte choisi dans les paramètres et son
identification Git Credential Manager. Il liste les dépôts personnels et ceux des
organisations accessibles, les branches et permet de créer un dépôt personnel ou
d’organisation. La création utilise le choix Privé/Public visible dans le formulaire.
Les droits du compte GitHub continuent de s’appliquer.

L’onglet **Pull requests** liste les PR ouvertes et fermées, leurs descriptions,
fichiers, commentaires et contrôles. Double-cliquer sur un fichier VBA ou un commentaire
de ligne ouvre le module local correspondant ; ce code local peut différer de celui de
la PR. La création utilise la branche active comme source et une branche cible choisie.
La branche doit déjà être publiée. Le bouton crée un brouillon par défaut.

Les outils de l’agent incluent `git_commit_selected`, `git_commit_read`,
`git_module_restore`, `git_pull_requests` et `git_pr_prepare`. Ce dernier conserve
un titre et une description dans le cache local : **Charger le brouillon préparé**
les insère dans le formulaire avant création. Les mutations locales conservent le
contrôle de révision `ExpectedState` et la politique d’édition configurée.

Les requêtes API sont envoyées uniquement à `api.github.com`, sans redirection HTTP.
Les identifiants ne sont ni sauvegardés dans le dépôt ni inclus dans les erreurs.
Les listes sont paginées. Il n’y a pas de nouvelle tentative automatique d’une création.
Fetch, la connexion et les consultations GitHub peuvent être annulés. Un import ou
une publication commencé ne propose pas d’interruption susceptible de laisser un état
ambigu ; consulter GitHub avant de renouveler une publication après une coupure réseau.

Référence de protocole : [API GitHub des pull requests](https://docs.github.com/en/rest/pulls/pulls).

## Interface et déploiement

**Paramètres > Apparence** propose Système, Clair et Sombre. La préférence est locale
à VBAi ; elle ne modifie pas le thème global du processus Office/SOLIDWORKS.
Les contrôles fixes restent dans les fichiers Designer. Les vues dynamiques du chat
utilisent WPF ; leur diff partage le même calcul que la vue Git WinForms.

Le Markdown est analysé avec Markdig : titres, listes, citations, liens, tableaux,
emphase et blocs de code. Le code VBA est coloré. Le HTML reste inerte et les images
distantes sont présentées comme liens, sans téléchargement automatique.
Les conversations restaurées chargent d’abord les 80 derniers messages puis proposent
de remonter dans l’historique. Les vues sont virtualisées ; toutes les données de
conversation restent disponibles pour la sauvegarde et l’agent.

Déployer le dossier complet de compilation, notamment Markdig et ses dépendances
System.Memory, System.Buffers, System.Numerics.Vectors et System.Runtime.CompilerServices.Unsafe.
Ne pas mélanger ce dossier avec la sortie du lanceur MSTest, qui utilise ses propres
versions de dépendances. L’installateur vérifie leur présence.

Validation : tests Git locaux avec VBE simulé, tests HTTP GitHub simulés,
chargement du designer, captures des thèmes et test de virtualisation. Les opérations
GitHub authentifiées et le rendu dans Excel/SOLIDWORKS réels nécessitent encore une
validation dans ces hôtes. Les nouveaux libellés sont présents dans les 13 catalogues ;
les traductions supplémentaires sont produites localement et restent à relire par
des locuteurs natifs.
