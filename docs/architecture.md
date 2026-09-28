# Organisation de la solution

Ouvrir `CodexVBE.sln` dans Visual Studio. Les cinq projets C# ciblent .NET Framework 4.8, C# 7.3 et x64 en Debug comme en Release. `Directory.Build.props` centralise ces paramètres et `.editorconfig` définit UTF-8, CRLF et les règles d’indentation.

## Projets et dépendances

| Projet | Responsabilité | Références de projet |
| --- | --- | --- |
| `src/CodexVBE/CodexVBE.csproj` | Complément COM chargé par Excel ou SOLIDWORKS. | Aucune. |
| `src/VBAi.Updater/VBAi.Updater.csproj` | Programme de mise à jour externe, copié dans la livraison du complément. | Aucune. |
| `tests/CodexVBE.Tests/CodexVBE.Tests.csproj` | Tests MSTest/VSTest unitaires, locaux et hôtes opt-in. | Complément et simulation des fournisseurs. |
| `tests/CodexVBE.Git.Smoke/CodexVBE.Git.Smoke.csproj` | Diagnostic autonome Git et interface associée. | Complément. |
| `tests/CodexVBE.Providers.Smoke/CodexVBE.Providers.Smoke.csproj` | Diagnostic des protocoles et processus CLI simulé. | Complément. |

Les trois projets de test référencent la DLL de production ; aucun ne recompile ses sources. `InternalsVisibleTo` leur donne accès aux membres internes. Les scénarios des diagnostics sont liés dans le projet MSTest sous `Shared`, pour tester la même DLL dans le processus mesuré.

Les diagnostics conservent les noms `GitTests.exe` et `ProviderTests.exe`. Le second simule aussi un processus Copilot ; sa copie vers la sortie VSTest utilise le chemin résolu par MSBuild, y compris en compilation isolée.

## Code du complément

Les chemins suivants sont relatifs à `src/CodexVBE/`.

| Dossier | Responsabilité |
| --- | --- |
| `Host` | Cycle de vie COM, menus VBE, volet latéral et journal. |
| `Bridge` | Protocole local et serveur de commandes dans l’hôte. |
| `Vbe` | Session VBE et résolution du projet ciblé. |
| `Vbe/Code` | Navigation et modification des modules et procédures VBA. |
| `Vbe/Debug` | Exécution, commandes natives du débogueur et lecture de ses fenêtres. |
| `Vbe/Forms` | UserForms, contrôles, propriétés, conteneurs et événements. |
| `Vbe/Projects` | Composants et références des projets. |
| `Vbe/Windows` | Fenêtres et volets de code VBIDE. |
| `Llm/Chat` | Conversations, contexte VBE et outils utilisables par les modèles. |
| `Llm/Providers` | Protocoles, authentification et catalogues de modèles. |
| `Llm/Settings` | Persistance des paramètres et formulaire de configuration. |
| `Llm/Settings/Views` | Vues WinForms Fournisseur, Compte GitHub et Apparence. |
| `Llm/Controls` | Contrôles nécessaires au concepteur du chat. |
| `Llm/Controls/Transcript` | Vues WinForms éditables des messages, activités, cartes, suggestions et contenu Markdown. |
| `Git` | Dépôts VBA, snapshots, synchronisation et interface GitHub. |
| `Git/Views` | Vues WinForms des onglets Git et GitHub. |
| `Editor` | Éditeur Monaco, synchronisation des modules, brouillons et navigation. |
| `Ui` | Contrôles partagés, thèmes WinForms et thème VBE natif expérimental, Markdown et comparaison de code. |
| `Localization` | Catalogues de traduction et résolution des textes. |
| `Properties` | Identité de l’assembly et visibilité accordée aux tests. |

Les fichiers `.cs`, `.Designer.cs` et `.resx` restent réunis avec leurs métadonnées `SubType` et `DependentUpon` pour le concepteur WinForms. Les icônes restent sous `assets/icons/` avec leurs noms de ressources embarquées inchangés.

Les tests unitaires suivent les dossiers et fichiers sous `tests/CodexVBE.Tests/Unit/`, avec un suffixe `.Tests.cs` par fichier source, y compris les classes partielles. Les parcours transversaux sont sous `Scenarios/`, les intégrations avec stockage ou hôte sous `Integration/`, les doubles partagés sous `Infrastructure/Fixtures/`, et les fixtures Excel et le client de passerelle sous `Infrastructure/Hosts/`. Les sondes PowerShell restent sous `tools/`, avec des liens dans VSTest pour les diagnostics manuels. Voir [la convention miroir](../tests/README.md#convention-miroir).

## Moteur de thème natif

`src/CodexVBE.Native/` contient le moteur C++ x64 des barres VBE. La solution expose ses sources dans un dossier ; le projet du complément déclenche `tools/build/Build-NativeRenderer.ps1` avant sa compilation et embarque la DLL produite. Le développeur doit disposer des outils C++ x64 Visual Studio et du SDK Windows ; le script respecte le dossier SDK déclaré dans le registre, y compris sur un autre disque. L'utilisateur final n'a pas besoin du compilateur ou d'un runtime C++ séparé.

Les tests C# du chargeur et du thème suivent la convention miroir. `tests/native/` contient le programme de cycle C++ et la vérification de chargement ; sa fixture `VBE7.dll` est synthétique et reste dans les artefacts de test. Ces vérifications ne remplacent ni une mesure de couverture C++ ni les essais visuels Excel/SOLIDWORKS. Voir le [thème natif](native-dark-theme.md) et les [commandes des tests natifs](../tests/native/README.md).

## Compilation et installation

Depuis la racine du dépôt :

```powershell
dotnet build CodexVBE.sln -c Debug -p:Platform=x64
dotnet build CodexVBE.sln -c Release -p:Platform=x64
```

La sortie habituelle du complément reste `bin/<Configuration>/net48/CodexVBE.dll`, pour préserver le chemin `CodeBase` de l’installation. Les sorties des tests restent dans leurs propres dossiers `bin/`. Les fichiers intermédiaires sont propres à chaque projet dans `obj/`.

Pour compiler sans écraser une DLL chargée par un hôte, fournir un chemin absolu :

```powershell
dotnet build CodexVBE.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build"
```

Chaque projet dispose alors de sa sortie `artifacts/build/<NomProjet>/<Configuration>/net48/`. Éviter un `OutputPath` partagé entre projets. Voir [les tests](../tests/README.md) pour les scénarios hôtes opt-in.

## Identité COM et frontières

Le nom `CodexVBE.dll`, le namespace `CodexVBE`, le ProgID, les GUID COM et les noms des ressources restent stables. La règle IDE0130 est désactivée sous `src` pour conserver ce namespace à travers les dossiers fonctionnels.

Le complément reste une seule assembly : `VbeSession`, les outils LLM, les interfaces et les objets COM partagent encore des contrats internes. Extraire ces couches en bibliothèques exige de définir leurs interfaces puis de vérifier le déploiement des dépendances dans les deux hôtes. Cette organisation sépare les responsabilités et les tests sans changer ces contrats lors du déplacement des fichiers.

## Programme de mise à jour

`src/VBAi.Updater` produit une application WinForms .NET Framework 4.8 x64.
`src/CodexVBE/Updates` contient le flux GitHub, les préférences, la coordination,
les travaux persistants et leurs interfaces. Le projet externe lie les seules
sources nécessaires et les catalogues ; il ne référence pas `CodexVBE.dll`, afin
que l’installeur puisse remplacer le complément après fermeture des hôtes.
La compilation du complément copie `VBAi.Updater.exe` et sa configuration dans
son dossier de sortie. Le protocole complet figure dans [updates.md](updates.md).

## Documentation IntelliSense

Les commentaires XML couvrent les déclarations publiques et privées du complément : classes, fonctions, constructeurs, propriétés, champs, événements, délégués et valeurs d’énumération. Le 29 septembre 2026, l’audit des 294 sources compte **5 208 déclarations documentées sur 5 208**, sans commentaire invalide ni erreur syntaxique. Une comparaison Roslyn avec le code précédant le lot documentaire constate zéro différence hors commentaires et espaces.

Pour conserver ce contrôle après une modification :

```powershell
dotnet run --project tools/XmlDocumentationAudit -- src/CodexVBE
```

Le mode `--compare <dossier-source-de-référence>` vérifie aussi l’équivalence syntaxique. Ces compteurs concernent le complément ; ils ne mesurent ni la couverture des tests ni la documentation de leurs fixtures.
