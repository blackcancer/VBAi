# À propos de VBAi

## Accès

- **Outils → À propos de VBAi** dans le VBE, même si le chat est fermé.
- **Menu ⋯ du chat → À propos de VBAi**.

Le dialogue affiche le logo VBAi, la signature « Your AI agent for VBA », une
présentation du produit et la version de l’assembly installé. La version informative
est utilisée si elle est disponible ; sinon la version de l’assembly est affichée.
Les autres champs présentent l’hôte actuel, Windows, l’architecture du processus,
.NET Framework 4.8 et la langue de l’interface. Excel et SOLIDWORKS sont identifiés
par leur nom de processus ; dans un essai autonome, le nom du processus de test
est affiché.

## Ressources et support

Les liens ouvrent le dépôt GitHub existant, sa documentation et la création d’un
signalement dans le navigateur par défaut. Ils ne sont ouverts qu’au clic.

**Copier les détails techniques** place dans le presse-papiers : version VBAi,
nom du processus hôte, plateforme et architecture, version CLR, culture de
l’interface et choix de thème. Aucun chemin de document, code VBA, message,
paramètre fournisseur ou identifiant de compte n’est consulté pour construire ce texte.
Les échecs du navigateur ou du presse-papiers sont signalés dans la fenêtre.

Fermer est le bouton par défaut ; Échap ferme le dialogue. La disposition et les tooltips sont définis
dans le Designer. Les 13 catalogues existants contiennent les nouveaux libellés,
rédigés localement. La signature du produit reste en anglais. Le thème et le sens
de lecture suivent les réglages existants ; les valeurs techniques restent
orientées de gauche à droite.

## Ressources WinForms

Le logo utilise une ressource Bitmap WinForms standard. Le projet active les
ressources présérialisées et déploie `System.Resources.Extensions.dll` avec les
autres dépendances. Ce lecteur Microsoft prend en charge les ressources objets :
[System.Resources.Extensions](https://www.nuget.org/packages/System.Resources.Extensions/8.0.0).

Le chargement de la fenêtre résout temporairement l’identité historique du lecteur
inscrite par le SDK vers la dépendance livrée. Le gestionnaire est retiré dans un
`finally` après `InitializeComponent` ; aucune configuration d’Excel, SOLIDWORKS
ou Visual Studio n’est modifiée. L’installateur vérifie la présence de cette DLL.

## Validation locale

- Compilation solution : sans erreur ni avertissement.
- DesignSurface : 28 surfaces et 282 contrôles enfants, avec édition et sérialisation.
- Suite unitaire complète : **1 139 tests réussis** ; **16 tests ciblés** rejoués
  après les derniers ajustements de contraste et de libération des images.
- Tests dédiés : métadonnées dans les 13 langues, liens, copie, erreurs, constructeur
  Designer, accès depuis le chat et le VBE simulé, cycle de vie du menu.
- Captures natives autonomes inspectées en français clair, anglais sombre et arabe
  sombre ; contrôle des débordements des contrôles. Aucun navigateur ni macro
  n’est lancé par ces validations.

Preuves : `artifacts/about/designer/`, `artifacts/about/tests/` et
`artifacts/about/screens/`. Ces essais ne remplacent pas une ouverture manuelle
dans le VBE d’Excel/SOLIDWORKS ni un enregistrement du source dans Visual Studio.
