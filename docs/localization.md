# Langues de l’interface VBAi

VBAi propose 13 variantes d’interface. La sélection est automatique au
chargement du complément :

1. Les légendes des menus natifs du VBE déterminent la langue de l’IDE :
   les légendes traduites de `Affichage` / `Outils` sont reconnues, y compris les
   suffixes d’accélérateurs des menus asiatiques (`表示(&V)`).
2. Si les menus ne sont pas accessibles, VBAi utilise la langue d’affichage Windows
   (`CurrentUICulture`), avec les correspondances régionales ci-dessous.
3. Une langue non prise en charge utilise l’anglais.

Un changement de langue de l’IDE prend effet au prochain chargement du complément.
La culture du thread hôte et les formats numériques et de dates de VBA ne sont pas modifiés.
La correction orthographique du composeur reçoit la culture retenue ; la disponibilité
effective des dictionnaires dépend de Windows/WPF.

| Langue | Culture de l’interface | Variantes |
|---|---|---|
| Anglais | en-US | Famille en |
| Français | fr-FR | Famille fr, dont fr-CA |
| Espagnol | es-ES | Famille es, dont es-MX |
| Allemand | de-DE | Famille de |
| Portugais brésilien | pt-BR | Famille pt, dont pt-PT |
| Italien | it-IT | Famille it |
| Japonais | ja-JP | Famille ja |
| Coréen | ko-KR | Famille ko |
| Chinois simplifié | zh-CN | zh-Hans, zh-SG et zh par défaut |
| Chinois traditionnel | zh-TW | zh-Hant, zh-HK, zh-MO |
| Russe | ru-RU | Famille ru |
| Arabe | ar-SA | Famille ar |
| Hindi | hi-IN | Famille hi |

Ce choix est une couverture pratique de langues largement utilisées, pas un classement
statistique des installations Windows. Pour l’arabe, les fenêtres et le chat utilisent
le sens de lecture droite à gauche ; les URL, chemins, sources VBA et diffs restent
de gauche à droite.

## Contenu couvert

- Fenêtres de chat, paramètres, GitHub et validation des modifications, menus VBE et menu du chat.
- Infobulles, accessibilité, statuts, messages de connexion et opérations Git.
- Cartes de conversation, diffs, annulation, historique, mémoire et références.
- Commandes `/explain`, `/fix`, `/refactor`, `/document`, `/tests`, `/plan` en anglais ;
  les commandes françaises existantes restent acceptées dans toutes les langues.
  Les descriptions sont traduites ; les identifiants des commandes restent stables.

Le code, les noms de projets/modules, les titres saisis et les messages existants
ne sont pas traduits. Les réponses du modèle et les diagnostics externes conservent
la langue de leur source. Les identifiants persistants de fournisseurs, les rôles
de messages et les chemins de données restent compatibles avec les sessions existantes.

## Maintenance

`src/VBAi/Localization/UiStrings.resx` contient les clés anglaises et
`UiStringsFrench.resx` leurs traductions françaises. Les autres catalogues suivent
le même schéma et sont déclarés dans `UiLanguages.cs`. Tous les catalogues sont
embarqués dans la DLL principale ; aucun assembly satellite à déployer.

Les contrôles fixes restent construits dans les fichiers `Designer.cs`, avec des
textes anglais éditables. `UiText.Apply` applique leurs traductions après
`InitializeComponent`, avant tout chargement de données utilisateur. Le designer
reste en anglais ; la traduction ne construit aucun contrôle. Les textes dynamiques
utilisent `UiText.Get` avec une clé anglaise. Ajouter chaque nouvelle clé à tous les
catalogues, y compris ses espaces et paramètres de format éventuels.

Compiler avec Visual Studio ou `dotnet build`, puis lancer sous Windows PowerShell :

```powershell
powershell.exe -Sta -NoProfile -File tools/tests/Test-Localization.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatDesigner.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-MultilingualWindows.ps1 -AssemblyPath artifacts/multilingual-tests/VBAi.dll
```

Les tests couvrent les 13 variantes sur les quatre fenêtres, la priorité VBE/Windows,
le repli anglais, la parité des ressources, les identifiants persistants et les alias
de commandes. Les captures multilingues sont générées dans `artifacts/localization/multilingual-windows`.
Ces tests utilisent un VBE simulé ; la détection dans Excel et SOLIDWORKS réels
reste à valider après chargement du complément.

## Production des nouveaux catalogues

Les onze catalogues supplémentaires ont été générés **sur cet ordinateur**, avec
des modèles Argos et CTranslate2, puis corrigés pour les principaux libellés et les
messages sensibles. Aucun texte n’a été envoyé à un service de traduction.
La variante traditionnelle utilise OpenCC `s2twp` et des corrections propres à Taïwan.
Une relecture linguistique native complète reste à faire ; les tests techniques ne
garantissent pas la qualité de toutes les formulations.

`tools/localization/translate_offline.py` ne contient aucun accès réseau. Il utilise
les modèles déjà présents dans `artifacts/localization/models`, préserve les termes
techniques et applique `tools/localization/overrides.json`. Les modèles et les dépendances
Python ne font pas partie du produit ni de son installation. Voir le README du dossier
`tools/localization` pour reproduire cette étape.
