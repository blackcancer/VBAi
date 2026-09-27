# GitHub et projets VBA

## Utilisation

Dans le menu `…` du chat, ouvrir **GitHub · synchroniser le VBA…** pour le document sélectionné.
Le document doit être enregistré, son projet déverrouillé et le VBE en mode conception.

1. Saisir l’URL HTTPS d’un dépôt GitHub existant et la branche, puis **Lier le dépôt**.
   Cette action consulte la branche distante, sans publier ni importer de code.
2. **Comparer** affiche les fichiers ajoutés, modifiés et supprimés depuis le dernier commit/import.
   Sélectionner un fichier pour ouvrir la comparaison avant/après avec numéros de lignes et couleurs.
3. Saisir un message puis **Commit (tout)** : export du VBA vivant et commit local de tous les changements.
4. **Push** publie uniquement les commits locaux sur la branche sélectionnée.
5. **Fetch** actualise les informations distantes sans toucher au VBA.
6. **Pull et importer** télécharge, vérifie puis importe les sources avec sauvegarde préalable.
7. **Restaurer VBA** restaure l’état précédant le dernier import si le code n’a pas changé depuis sa relecture.

L’organisation reprend les vues **Modifications Git** et **Historique** de Visual Studio : commit local séparé
du push, fetch sans modification du projet, compteurs entrants/sortants et comparaison des fichiers.
Les compteurs reflètent le dernier fetch ; l’historique montre les 40 derniers commits de la branche locale.
L’aperçu du diff est limité à 2 000 lignes ; les sources commitées restent complètes.

## Prérequis et authentification

Dans **Paramètres…**, la section **GitHub · dépôts** est disponible quel que soit le fournisseur IA.
**Se connecter à GitHub** lance la connexion par navigateur de Git Credential Manager.
**Actualiser les comptes** relit les comptes mémorisés, sans demander de jeton ni ouvrir de connexion.
Choisir le compte puis **Enregistrer** conserve uniquement son identifiant dans les paramètres de CodexVBA.
Le choix s’applique aux prochaines fenêtres de synchronisation ; les commandes Git ciblent ce compte
avec des options locales au processus, sans modifier la configuration Git globale.
Le choix automatique laisse Git décider. Un compte mémorisé ne prouve pas les droits d’accès au dépôt :
ceux-ci sont vérifiés lors des échanges réseau.

La connexion au navigateur prend effet immédiatement dans Git Credential Manager, même si le formulaire
est ensuite annulé. Annuler ne modifie pas le compte sélectionné dans les paramètres de CodexVBA.
L’authentification du fournisseur IA Copilot reste distincte.

- Git for Windows (`git.exe` dans PATH), identité `user.name` / `user.email` configurée.
- Authentification HTTPS via le gestionnaire d’identifiants Git déjà configuré, par exemple Git Credential Manager.
- Aucun jeton dans le formulaire, les sources, le manifeste ou les paramètres de CodexVBA.
- La création du dépôt GitHub s’effectue actuellement sur GitHub. Les URL SSH et GitHub Enterprise ne sont pas encore proposées.

## Stockage privé

Aucun fichier ni dossier n’est ajouté à côté du classeur Excel ou de la macro SolidWorks.

`%LocalAppData%\CodexVBE\Git\<empreinte du chemin du document>\` contient :

- `binding.json` : URL et branche, sans secret ;
- un dépôt **bare** privé par liaison dépôt/branche, sans copie de travail ;
- un verrou empêchant deux fenêtres/processus de synchroniser simultanément ce document.

Git reçoit les blobs, arbres et messages de commit via ses entrées/sorties standard.
Les sources sont assemblées et comparées en mémoire. Les API natives `VBComponent.Export` et
`VBComponents.Import` exigent cependant des fichiers : ceux-ci sont créés dans un sous-dossier unique
de `%LocalAppData%\CodexVBE\GitTemporary`, puis supprimés. Un arrêt brutal peut laisser ce dossier temporaire.

La liaison est **locale à l’ordinateur et au chemin enregistré du document**. Elle n’est pas embarquée dans
le classeur ou la macro. Après déplacement/renommage ou sur un autre ordinateur, relier le dépôt.
Les sauvegardes et commits non poussés doivent être conservés : ne pas purger ce cache comme un cache jetable.

## Format versionné

Seul le sous-arbre `vba/` est remplacé par un commit de sources. Les autres fichiers du dépôt sont conservés.
Ne pas placer de fichiers non gérés dans `vba/` : ils sont refusés lors de la lecture du manifeste.

| Fichier | Traitement |
| --- | --- |
| `manifest.json` | Version du format, noms/types des composants, présence de ressources, GUID/versions des références |
| `Nom.bas` | Module standard, export natif avec attributs |
| `Nom.cls` | Classe, export natif avec attributs |
| `Nom.frm` + `Nom.frx` | Formulaire et ressources binaires associées |
| `Nom.vba` | Code visible d’un module lié au document, par exemple `ThisWorkbook` ou un module de feuille |

Les textes sont normalisés en UTF-8 et LF dans Git. À l’import, ils sont convertis vers la page de codes
Windows du VBE ; une conversion impossible échoue avant modification du projet. Les `.frx` restent binaires.
Les conteneurs Excel/SolidWorks, données du classeur, conversations et réglages des fournisseurs ne sont pas exportés.

## Import et restauration

- Premier import : le bouton **Pull et importer** remplace explicitement les sources locales, après sauvegarde.
- Imports suivants : les changements locaux non commitées bloquent le pull. Les historiques divergents sont refusés,
  sans merge automatique, rebase ni push forcé.
- Les modules hôtes doivent déjà exister avec les mêmes noms ; leur code est remplacé sur place.
  CodexVBA ne crée ni ne supprime des feuilles ou des objets du document.
- Les modules/classes/formulaires modifiés sont réimportés ; ceux qui ont disparu du manifeste sont supprimés.
- Les références doivent correspondre. Aucun changement automatique des bibliothèques installées.
- Les sources sont relues après import. Une erreur COM peut survenir après application : aucun import n’est relancé automatiquement.
- Avant toute mutation, une sauvegarde Git et un marqueur de récupération sont persistés. Les références privées
  `refs/codex/*` ne sont jamais poussées.
- Une erreur laisse la sauvegarde et bloque les synchronisations suivantes jusqu’à restauration.
  Si la relecture de l’état partiel échoue elle aussi, la restauration automatique est bloquée ; la sauvegarde reste accessible
  dans `refs/codex/backup` pour une récupération manuelle.
- Une restauration réussie laisse les commits distants intacts. Après un import réussi, elle apparaît comme une modification
  locale pouvant être publiée par un nouveau commit ; un pull ne l’écrase pas silencieusement.
- Un pull sans changement du VBA conserve la sauvegarde précédente.

Après import/restauration, compiler/vérifier le projet puis enregistrer le document dans son application.
L’intégration ne lance aucune macro et n’enregistre pas automatiquement le document. Une modification de VBA signé
doit suivre le processus habituel de signature du projet.

## Périmètre actuel et validation

La première version n’inclut pas encore le staging par fichier, le changement de branche avec modifications locales,
un éditeur de fusion, les pull requests ou la création de dépôts depuis l’interface.

`tools/tests/Git/GitTests.csproj` teste de vrais dépôts Git locaux, les échanges push/fetch,
les protections contre les divergences, les sauvegardes privées, le cycle complet des commandes WinForms,
les erreurs COM simulées, les modules hôtes et les ressources de formulaires. Le formulaire est aussi chargé
dans `DesignSurface` avec un constructeur sans services actifs.

La validation automatisée utilise un VBE simulé. Les imports réels dans Excel et SolidWorks et les échanges
avec un dépôt GitHub authentifié restent **NOT_RUN**. Aucun code utilisateur n’a été publié pendant ces tests.

## Références techniques

- [Visual Studio : fetch, pull, push et sync](https://learn.microsoft.com/en-us/visualstudio/version-control/git-fetch-pull-sync)
- [VBIDE : Export](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/export-method-vba-add-in-object-model)
- [VBIDE : Import](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/import-method-vba-add-in-object-model)
- [Git : commit-tree](https://git-scm.com/docs/git-commit-tree)
- [Git : update-ref](https://git-scm.com/docs/git-update-ref)
- [Git Credential Manager : comptes multiples](https://github.com/git-ecosystem/git-credential-manager/blob/main/docs/multiple-users.md)
