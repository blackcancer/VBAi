# Tests de CodexVBE

Les trois projets de test et diagnostic sont regroupés ici et visibles dans `CodexVBE.sln`. Ils ciblent .NET Framework 4.8 x64 et référencent le véritable complément sous `src/CodexVBE`.

## Organisation

- `CodexVBE.Tests/Unit/` : miroir des dossiers **et des fichiers** de production, avec le suffixe `.Tests.cs`.
- `CodexVBE.Tests/Scenarios/` : scénarios complémentaires qui vérifient plusieurs fichiers ensemble.
- `CodexVBE.Tests/Infrastructure/Fixtures/` : doubles de test, utilitaires et initialisation partagés.
- `CodexVBE.Tests/Integration/` : stockage local, Git, fournisseurs et hôtes Excel/SOLIDWORKS.
- `CodexVBE.Tests/Infrastructure/Hosts/` : fixtures hôtes et client du tube nommé.
- `CodexVBE.Git.Smoke/` : scénarios Git partagés avec VSTest et exécutable `GitTests.exe`.
- `CodexVBE.Providers.Smoke/` : scénarios fournisseurs partagés et exécutable `ProviderTests.exe`.

Les sources partagées apparaissent sous `Shared` dans VSTest. Les scripts PowerShell restent sous `tools/tests` et `tools/probes`, avec des liens sous `Manual/Debug` dans Visual Studio. Ils ne sont ni des tests MSTest ni du code livré dans le complément.

### Convention miroir

| Production sous `src/CodexVBE/` | Tests sous `CodexVBE.Tests/Unit/` |
| --- | --- |
| `Host/AddIn.cs` | `Host/AddIn.Tests.cs` |
| `Llm/Chat/ChatWindow.Sessions.cs` | `Llm/Chat/ChatWindow.Sessions.Tests.cs` |
| `Vbe/Forms/VbeForms.CheckBoxDuplication.cs` | `Vbe/Forms/VbeForms.CheckBoxDuplication.Tests.cs` |

Un nouveau test ciblé rejoint le fichier miroir de l’implémentation, y compris pour une classe partielle.
Les parcours qui vérifient plusieurs surfaces ensemble restent dans `Scenarios` ; les parcours avec
stockage ou hôte réel restent dans `Integration`. Les catégories et identifiants VSTest existants
sont conservés. Les classes de test utilisent `partial` pour partager leurs auxiliaires sous
`Infrastructure/Fixtures`, sans recopier les doubles COM ou renommer leurs types utilisés par réflexion.

Il n’y a pas de fichier miroir vide pour simuler une couverture. L’absence de miroir dédié ne prouve
pas l’absence de couverture par un scénario ; seule la mesure de couverture établit les lignes et
branches exécutées. Pour contrôler la convention et obtenir l’inventaire des correspondances :

```powershell
powershell.exe -NoProfile -File tools/tests/Test-TestLayout.ps1 -ReportPath artifacts/test-layout/mirror-inventory.json
```

## Suite locale

Depuis la racine du dépôt :

```powershell
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj -c Debug
```

Pour une compilation isolée, notamment quand Excel a chargé la DLL installée :

```powershell
dotnet build CodexVBE.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --results-directory artifacts/test-results --logger "trx;LogFileName=tests.trx"
```

MSBuild copie automatiquement le processus CLI simulé dans la sortie VSTest. Les tests fournisseurs utilisent des réponses HTTP et des processus simulés ; les essais authentifiés ne sont pas lancés par défaut. Les tests WinForms nécessitent Windows et une session interactive. Les tests Git nécessitent `git.exe` et travaillent sur des dépôts temporaires locaux.

Le filtre `--filter TestCategory=Unit` limite l’exécution aux tests unitaires. Une validation globale doit exécuter la suite complète sans ce filtre. Les résultats TRX font foi pour le nombre de tests exécutés, échoués et ignorés ; les anciens pourcentages ne sont pas des résultats actuels.

## Couverture

```powershell
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj -c Debug --collect:"XPlat Code Coverage" --results-directory artifacts/coverage -- DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Format=cobertura DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Exclude=[ProviderTests]*
```

Le filtre exclut uniquement l’exécutable auxiliaire des mesures du complément. Aucun fichier, méthode ou branche de production ne doit être exclu pour atteindre la cible. Le code exécuté dans Excel ou SOLIDWORKS n’est pas mesuré par le collecteur du processus VSTest.

## Excel automatisé — hôte prioritaire

```powershell
$env:CODEXVBE_RUN_EXCEL_TESTS = '1'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=Excel
Remove-Item Env:CODEXVBE_RUN_EXCEL_TESTS
```

`ExcelVbeFixture` crée une instance Excel visible et un classeur temporaire, ouvre le VBE par `CommandBars.ExecuteMso("VisualBasic")`, puis appelle le tube `CodexVBE.<PID>`. Les tests vérifient le chargement du complément, le projet ciblé et sa sauvegarde. Ils ferment uniquement leur propre classeur et processus. L’installation du complément est un prérequis ; les tests ne changent pas AccessVBOM. Une compilation isolée ne remplace pas la DLL installée.

### Services de langage Monaco dans Excel

La matrice JavaScript est intégrée à VSTest par `MonacoLanguageScriptScenarios` et nécessite Node.js. Elle vérifie les références dynamiques, les membres, le survol, les signatures et les règles de blocs et de formatage :

```powershell
node --test tools/tests/Test-MonacoLanguage.mjs tests/CodexVBE.Tests/Infrastructure/Fixtures/Editor/MonacoEditing.Scenarios.mjs
```

Le parcours `MonacoLanguageExcelTests` ouvre Excel et le véritable WebView Monaco, vérifie les objets et collections Excel, les fonctions et alias VBA (`Left`/`Left$`), ajoute puis retire les références Office et Scripting, vérifie le survol, crée des blocs, annule et formate. Sa minuterie de synchronisation est arrêtée après l’initialisation pour vérifier que les services de langage ne modifient pas le module natif. Aucun fichier utilisateur ni macro n’est exécuté. Il refuse de démarrer si une session Excel existe déjà.

```powershell
$env:VBAI_EDITOR_LANGUAGE_EXCEL_TEST = '1'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter FullyQualifiedName~MonacoLanguageExcelTests
Remove-Item Env:VBAI_EDITOR_LANGUAGE_EXCEL_TEST
```

## SOLIDWORKS préchargé

```powershell
$env:CODEXVBE_SOLIDWORKS_PID = '<PID SLDWORKS existant>'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=SolidWorks
Remove-Item Env:CODEXVBE_SOLIDWORKS_PID
```

L’utilisateur doit avoir ouvert SOLIDWORKS, son VBE et le complément. Le test utilise la passerelle du PID fourni, vérifie le projet et le complément connecté, sans lancer ni fermer SOLIDWORKS. Sans activation explicite, les tests hôtes sont ignorés : ils restent `NOT_RUN`, pas validés par les simulations locales.
