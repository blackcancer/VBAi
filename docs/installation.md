# Installation et diagnostic

Le complément cible Windows et un hôte **VBE 64 bits**. Ouvrir `VBAi.sln` dans Visual Studio, avec Debug ou Release et la plateforme x64. Les scripts d’installation actuels attendent la sortie **Debug** sous `bin/Debug/net48/`.

## Compiler et inscrire

Fermer les hôtes ayant chargé la DLL avant de reconstruire leur sortie installée. Pour continuer à compiler pendant qu’un hôte est ouvert, utiliser la sortie isolée décrite dans [Architecture](architecture.md).

La compilation utilise aussi la PIA `Microsoft.Office.Interop.PowerPoint` 15.0 pour lire le handle PowerPoint via son interface COM. Le projet la recherche dans les outils Office de Visual Studio puis dans le GAC ; un autre emplacement peut être fourni avec `-p:PowerPointInteropPath="chemin/Microsoft.Office.Interop.PowerPoint.dll"`. `EmbedInteropTypes=true` embarque les types nécessaires : cette PIA n'est pas une dépendance à livrer avec le complément.

Depuis un PowerShell de développement Visual Studio **64 bits**, à la racine du dépôt :

```powershell
dotnet build VBAi.sln -c Debug -p:Platform=x64
TlbExp.exe .\bin\Debug\net48\VBAi.dll /out:.\bin\Debug\net48\VBAi.tlb
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Install-VBAi.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Test-VBAiInstallation.ps1
```

`TlbExp.exe` provient du SDK .NET Framework installé avec les outils de développement. Garder toute la sortie de compilation, notamment les dépendances de rendu du chat. L’installeur refuse une bibliothèque de types absente ou plus ancienne que la DLL et vérifie l’identité COM attendue.

Le produit conserve le ProgID `VBAi.AddIn` et le contrôle COM `VBAi.ChatToolWindow`. L’inscription du complément est utilisateur ; le volet natif nécessite aussi une résolution de son ProgID dans la vue machine 64 bits de l’hôte testé.

## Registre et environnement Codex

Depuis un shell Codex identifié par `CODEX_SHELL=1`, les scripts passent par `Invoke-VBAi-OutsideSandbox.ps1` et une tâche planifiée temporaire. Cette voie écrit et vérifie la ruche Windows réelle, avec l’élévation nécessaire, puis supprime la tâche. Ne pas conclure qu’une inscription est correcte sur la seule lecture de la vue isolée du shell.

Le chemin de découverte VBE qualifié sur ce poste est `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64`. L’installeur vérifie aussi le ProgID, le CLSID, le CodeBase, la bibliothèque de types et `LoadBehavior`.

## Ouvrir et vérifier

Dans Excel, ouvrir le VBE depuis la commande de l’application. Pour SOLIDWORKS, ouvrir une macro et son IDE manuellement. Le complément ajoute **Affichage → Assistant VBAi**, **Affichage → GitHub VBAi…** et **Outils → Configuration VBAi…**.

Pour inspecter le pont d’un PID connu :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-VBAi.ps1 -HostProcessId 12345 -Command status
```

Remplacer `12345` par le PID de l’hôte attendu. `status` donne le chemin de l’assembly chargée, son identifiant de compilation et l’état de connexion. Après mise à jour, redémarrer entièrement l’hôte : fermer une fenêtre VBE peut seulement la masquer et laisser la DLL chargée.

Le VBE mémorise l’ancrage de son volet. Sur une disposition vierge, le premier placement à droite peut nécessiter un déplacement manuel ; la restauration à droite est qualifiée dans Excel.

## Désinstaller ou diagnostiquer un chargement

Fermer les hôtes VBE avant `tools/Uninstall-VBAi.ps1`. En cas de complément absent ou impossible à charger, commencer par `Test-VBAiInstallation.ps1`, le chemin de DLL retourné par `status` lorsqu’il répond, l’architecture 64 bits et l’état du Gestionnaire de compléments. Les anciens essais et la cause de l’inscription isolée sont conservés dans [l’historique d’exploration](archive/exploration/project.md).

## Mises à jour et livraison

Le mécanisme [de mise à jour par releases GitHub](updates.md) est distinct de
l’inscription de développement décrite ci-dessus. `tools/Prepare-Release.ps1`
prépare un payload versionné pour le futur installeur, avec la TLB et le programme
externe d’application. La version COM reste stable ; la propriété MSBuild
`ProductVersion` pilote la version de livraison. L’installation autonome exige
le marqueur du protocole et un installeur signé disponible dans une release.
