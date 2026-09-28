# VBAi

**Your AI agent for VBA**

Complément COM pour le **Visual Basic Editor 64 bits**, notamment Excel et SOLIDWORKS, ciblant **.NET Framework 4.8 x64**. Il intègre une conversation LLM et des outils de lecture, d’édition, de conception de UserForms et de débogage du projet VBA vivant. Codex est le fournisseur par défaut ; les autres fournisseurs se configurent dans le complément.

Anciennement CodexVBE : les noms de solution, d’assembly, les identifiants COM et les répertoires de données sont conservés pour les installations existantes. L’interface dispose de 13 variantes linguistiques.

## Documentation

- [Index de la documentation](docs/README.md)
- [Installation et diagnostic](docs/installation.md)
- [Conversation et sessions](docs/chat-ui.md)
- [Fournisseurs](docs/providers.md) et [GitHub](docs/github-integration.md)
- [Architecture](docs/architecture.md) et [concepteurs WinForms](docs/winforms-designer.md)
- [Catalogue des 177 outils LLM](docs/reference/vbe-tools.md)
- [Travaux restants](docs/roadmap.md) et [couverture des tests](docs/test-coverage.md)

## État vérifié

Après intégration des PR #2 et #3, mesure globale du code `fb166a4` : **1 151 tests réussis, aucun échec, 1 test SOLIDWORKS ignoré**, couverture de **100 % des lignes et des branches**. Les deux tests Excel passent. Les **24 concepteurs WinForms** restent une validation antérieure.

Les tests réels Excel vérifient le chargement, le pont et la sauvegarde/relecture d’un classeur macro jetable. Les essais historiques dans SOLIDWORKS 2019 SP5 sur `test.swp` vérifient compilation, exécution, breakpoint, pas à pas et reprise, avec restauration du code initial. Ces scénarios ne qualifient pas chaque fonction dans chaque hôte ; les preuves et limites figurent dans [l’état du projet](docs/project.md).

## Développement

Ouvrir `CodexVBE.sln` dans Visual Studio, plateforme **x64** :

```powershell
dotnet build CodexVBE.sln -c Debug -p:Platform=x64
```

`src/CodexVBE/` contient le complément ; `tests/` contient le projet VSTest, ses miroirs et les diagnostics. `tools/` regroupe installation, contrôleur, sondes et essais natifs. `assets/` contient les ressources graphiques ; `artifacts/` contient les rapports locaux ignorés par Git.

Les essais autonomes utilisent une instance Excel visible et jetable. SOLIDWORKS et son VBE sont préchargés par l’utilisateur. Les raccourcis VBE et `SendKeys` sont proscrits ; les actions ciblent le projet et vérifient leurs révisions. Pour compiler pendant qu’un hôte charge la DLL, utiliser une sortie isolée plutôt que la sortie installée : voir [Architecture](docs/architecture.md).
