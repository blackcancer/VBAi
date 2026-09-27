# Langues de l’interface VBAi

VBAi propose une interface française et anglaise. La sélection est automatique au
chargement du complément :

1. Les légendes des menus natifs du VBE déterminent la langue de l’IDE :
   `Affichage` / `Outils` pour le français, `View` / `Tools` pour l’anglais.
2. Si les menus ne sont pas accessibles, VBAi utilise la langue d’affichage Windows
   (`CurrentUICulture`), y compris les variantes régionales françaises.
3. Une langue non prise en charge utilise l’anglais.

Un changement de langue de l’IDE prend effet au prochain chargement du complément.
La culture du thread hôte et les formats numériques et de dates de VBA ne sont pas modifiés.
La correction orthographique du composeur utilise `fr-FR` ou `en-US` selon la langue retenue.

## Contenu couvert

- Fenêtres de chat, paramètres, GitHub et validation des modifications, menus VBE et menu du chat.
- Infobulles, accessibilité, statuts, messages de connexion et opérations Git.
- Cartes de conversation, diffs, annulation, historique, mémoire et références.
- Commandes `/explain`, `/fix`, `/refactor`, `/document`, `/tests`, `/plan` en anglais ;
  les commandes françaises existantes restent acceptées dans les deux langues.

Le code, les noms de projets/modules, les titres saisis et les messages existants
ne sont pas traduits. Les réponses du modèle et les diagnostics externes conservent
la langue de leur source. Les identifiants persistants de fournisseurs, les rôles
de messages et les chemins de données restent compatibles avec les sessions existantes.

## Maintenance

`src/Localization/UiStrings.resx` contient les clés anglaises et
`UiStringsFrench.resx` leurs traductions françaises. Les deux catalogues sont
embarqués dans la DLL principale ; aucun assembly satellite à déployer.

Les contrôles fixes restent construits dans les fichiers `Designer.cs`, avec des
textes anglais éditables. `UiText.Apply` applique leurs traductions après
`InitializeComponent`, avant tout chargement de données utilisateur. Le designer
reste en anglais ; la traduction ne construit aucun contrôle. Les textes dynamiques
utilisent `UiText.Get` avec une clé anglaise. Ajouter chaque nouvelle clé aux deux
catalogues, y compris ses espaces et paramètres de format éventuels.

Compiler avec MSBuild de Visual Studio, puis lancer sous Windows PowerShell :

```powershell
powershell.exe -Sta -NoProfile -File tools/tests/Test-Localization.ps1
powershell.exe -Sta -NoProfile -File tools/tests/Test-ChatDesigner.ps1
```

Les tests couvrent les deux langues sur les quatre fenêtres, la priorité VBE/Windows,
le repli anglais, la parité des ressources, les identifiants persistants et les alias
de commandes. Les captures sont générées dans `artifacts/localization/screenshots`.
Ces tests utilisent un VBE simulé ; la détection dans Excel et SOLIDWORKS réels
reste à valider après chargement du complément.
