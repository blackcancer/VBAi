# Confidentialité des projets dans le chat

## Accès par défaut

Une conversation est liée à une macro. Les outils de lecture **et** d'écriture qui ciblent un projet sont contrôlés avant leur exécution : lire un module du classeur B depuis un chat lié au classeur A est refusé. Un projet simplement ouvert dans la même instance n'est pas une autorisation de transmission.

`list_projects`, `status` et le contexte vivant ne transmettent que les projets autorisés. `debug_state` ne retourne un emplacement que lorsque le volet de code appartient réellement au projet demandé, indépendamment de la sélection dans l'arbre des projets. Les références et pièces jointes de code sont contrôlées avant résolution ou transmission.

Les noms utilisés par les références `#` et `@` sont rapprochés des chemins de la macro uniquement si l'identité est unique dans l'inventaire vivant. Un nom ambigu ne permet pas d'étendre les autorisations. Les chemins absolus sont préférés pour les projets enregistrés.

## Autorisations particulières

Le menu des options du chat contient **Accès aux projets**. La fenêtre WinForms possède son Designer : seules les lignes des projets ouverts sont alimentées à l'exécution.

- Les cases de projets supplémentaires donnent une autorisation de **lecture seulement**. Les écritures restent limitées à la macro liée.
- Une case distincte autorise le **contexte partagé du VBE et du presse-papiers**. Les fenêtres du débogueur, l'arbre natif, le navigateur d'objets, les collections et fenêtres globales ne peuvent pas toujours être attribués à une macro. Ces outils sont refusés par défaut, même si plusieurs projets ont été autorisés en lecture.
- Cette autorisation de contexte partagé permet de transmettre des données non filtrées, provenant éventuellement d'autres projets ou du presse-papiers système. Elle ne donne pas une autorisation d'édition d'un autre projet.
- La lecture d'un fichier externe reste soumise à son parcours distinct : chemin fourni par l'utilisateur et confirmation de transmission.

Toute modification de ces autorisations démarre une conversation vierge dans la même portée. Une donnée déjà envoyée au fournisseur ne peut pas être retirée ; l'historique d'une ancienne session conserve ses autorisations d'origine. Une branche de conversation qui reprend l'historique hérite aussi de ses autorisations. Les autorisations sont persistées avec la session et ne sont jamais accordées par un appel du modèle.

Les sessions enregistrées avant cette politique sont identifiées à la lecture de SQLite. À leur ouverture, les messages destinés au fournisseur, le fil Codex et le contexte de reprise sont réinitialisés. Le transcript local reste visible avec un message indiquant cette transition. Les anciens contenus ne sont donc pas réexpédiés automatiquement ; cette migration ne retire pas les données déjà reçues par un fournisseur.

Les branches de conversation n'incluent que les messages postérieurs à cette transition. Un message ancien reste consultable et exportable localement, mais ne peut pas servir à réintroduire implicitement son contexte dans un nouveau fil fournisseur.

## Portée technique

La garde s'applique aux outils LLM, y compris les passerelles du catalogue. Elle ne transforme pas l'instance VBE en environnement isolé : le VBE, les événements de l'éditeur et le pont local gardent leurs fonctions pour l'hôte. Les accès à des fichiers explicitement fournis et le contexte partagé sont des exceptions volontaires à la portée de la macro, et sont présentés comme telles.

Un nouvel outil sans argument `Project` doit être classifié explicitement comme indépendant du contenu des projets pour être accessible sans autorisation de contexte partagé. Les opérations qui ont un argument de projet mais retournent un inventaire global, notamment la fermeture d'un projet autonome, exigent cette autorisation aussi.

## Vérifications

Les tests vérifient le refus de lecture B, les lectures autorisées sans autorisation d'écriture, le filtrage des inventaires, les noms ambigus, les outils globaux, les alias des références, le contexte actif d'un autre projet, les lectures Monaco et l'héritage des permissions lors d'une branche de conversation. Ces tests ne prouvent pas que des données déjà transmises peuvent être rappelées au fournisseur.
