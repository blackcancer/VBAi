# VBAi

**Your AI agent for VBA**

Interface disponible en 13 variantes linguistiques selon la langue du VBE, avec repli sur la langue
d’affichage Windows puis sur l’anglais. Voir [la localisation](docs/localization.md).

Anciennement CodexVBE. Les noms de solution, d’assembly, les identifiants COM et les répertoires de données restent inchangés pour préserver les installations existantes.

Complément COM expérimental pour le VBE 64 bits, notamment dans Excel et SOLIDWORKS, ciblant .NET Framework 4.8 x64. Les essais autonomes passent par Excel ; l'utilisateur ouvre lui-même l'IDE de SOLIDWORKS avant les essais dans cet hôte. Le pilotage et la lecture du VBE ne doivent employer aucun raccourci clavier ni dépendre de son focus.

## Organisation

- CodexVBE.sln : solution à ouvrir dans Visual Studio Community, configuration Debug ou Release, plateforme x64.
- CodexVBE.csproj : projet .NET Framework 4.8.
- src/ : code du complément chargé dans l'hôte du VBE, organisé par Host, Bridge, Vbe et Llm ; voir [l'architecture](docs/architecture.md).
- tools/ : scripts d'installation, de diagnostic et d'appel de la passerelle.
- tools/VbeController/ : recherches et inventaire des fenêtres, sans interaction clavier.
- test.swp : macro de travail jetable, maintenue ouverte dans le VBE.
- test-backups/ : copie de secours fournie par l'utilisateur et sauvegardes de test. Ne pas restaurer lors des essais ordinaires.
- bin/ et obj/ : sorties de compilation.

## État vérifié

Le complément expose un tube nommé CodexVBE.<PID hôte> accessible au compte Windows courant. Les commandes status, list_projects, list_modules, read_module et replace_lines ont fonctionné sur un projet .swp ouvert dans SOLIDWORKS. Une retouche temporaire de test1 dans test.swp a été relue puis annulée via le VBE ; l'empreinte SHA-256 du code final égalait celle du code initial. Le fichier .swp étant verrouillé pendant l'ouverture, cette vérification porte sur le code vivant du VBE, pas sur un nouvel enregistrement sur disque.

Le contrôleur externe fournit `state`, `windows`, `find_text`, `list_symbols`, `locals` et `immediate`. Les quatre premières commandes ont été testées sur le VBE ouvert dans SOLIDWORKS ; `list_symbols` a retrouvé `test1.main` ligne 2 dans le projet VBA `test` (nom actuel de test.swp). Les lectures `locals` et `immediate` ont été vérifiées dans Excel.

Dans Excel, `locals` lit les lignes de la fenêtre Variables locales par UI Automation `ValuePattern`. En mode Arrêt, la ligne de `probeValue` indiquait `1`. `immediate` lit la fenêtre Exécution par `TextPattern` et a renvoyé la sortie `2` de `Debug.Print`. Ces commandes n'utilisent ni raccourci ni coordonnées. L'évaluation d'une expression dans la fenêtre Exécution reste à concevoir.

Les commandes `debug_state`, `list_commands`, `select_code` et `invoke_debug` sélectionnent une ligne avec le modèle VBIDE puis invoquent un contrôle natif du VBE par `CommandBarControl.Execute`, avec vérification du projet, du mode et de l'empreinte du module. Dans Excel, elles ont permis un cycle complet : arrêt à la ligne 4, pas détaillé vers la ligne 5, poursuite, retrait du point d'arrêt et nouvelle exécution sans arrêt. Cette version n'a pas encore été validée dans SOLIDWORKS. La passerelle n'est pas encore exposée directement comme outil Codex.

## Utilisation

La DLL et sa bibliothèque de types COM se trouvent sous bin/Debug/net48/. Ne pas les reconstruire pendant une session VBE qui les a chargées. Ouvrir CodexVBE.sln dans Visual Studio Community, compiler en x64, générer `CodexVBE.tlb` avec l'outil Microsoft `TlbExp.exe`, puis relancer l'hôte pour charger la nouvelle version. La compilation ne prend que les sources de `src/` ; les sondes de `tools/` n'entrent pas dans le complément.

Depuis un PowerShell de développement Visual Studio 64 bits, dans la racine du projet :

    TlbExp.exe .\bin\Debug\net48\CodexVBE.dll /out:.\bin\Debug\net48\CodexVBE.tlb
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Install-CodexVBE.ps1
    $hostProcessId = (Get-Process EXCEL | Select-Object -First 1).Id
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Invoke-CodexVBE.ps1 -HostProcessId $hostProcessId -Command status

Contrôleur externe (lecture seule) :

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\VbeController\Invoke-VbeController.ps1 -HostProcessId $hostProcessId -Command state
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\VbeController\Invoke-VbeController.ps1 -HostProcessId $hostProcessId -Command windows
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\VbeController\Invoke-VbeController.ps1 -HostProcessId $hostProcessId -Command find_text -Project VBAProject -Query main
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\VbeController\Invoke-VbeController.ps1 -HostProcessId $hostProcessId -Command list_symbols -Project VBAProject

L'installation écrit uniquement l'inscription COM et la bibliothèque de types de l'utilisateur courant, ainsi que l'entrée VBE 64 bits. Depuis un shell Codex, `Install-CodexVBE.ps1` lance automatiquement une tâche planifiée temporaire avec élévation Windows pour écrire dans la ruche réellement utilisée par Excel et SOLIDWORKS, puis vérifie le ProgID, le CLSID, le chemin de la DLL, la bibliothèque de types et `LoadBehavior`. La tâche est supprimée à la fin. Dans un PowerShell Windows ordinaire, l'installation utilisateur reste directe. `Test-CodexVBEInstallation.ps1` vérifie aussi la ruche réelle depuis Codex. Pour désinstaller, fermer les hôtes VBE puis lancer `Uninstall-CodexVBE.ps1`, qui utilise la même voie depuis Codex.

La commande replace_lines nécessite l'empreinte SHA-256 renvoyée par read_module et refuse une modification hors du mode conception. Les identifiants de projet VBA peuvent différer du nom du fichier : le projet de test.swp s'appelle actuellement `test`.

## Débogage préparé dans le code source

`debug_state` renvoie le mode du projet et la sélection du volet de code ; `list_commands` relève les contrôles natifs du VBE. `select_code` utilise `CodePane.SetSelection` sans frappe clavier. `invoke_debug` exige le projet, le module, la ligne, l'empreinte du code, le mode attendu, l'action (`toggle_breakpoint`, `run`, `continue`, `step_into` ou `step_over`) ainsi que l'identifiant et l'intitulé exacts du contrôle natif relevé par `list_commands`. La commande refuse un volet de code différent de celui demandé. Le résultat est vérifié par un appel ultérieur à `debug_state` ; `invoke_debug` ne présume pas qu'un point d'arrêt a été posé.

La DLL actuelle a été recompilée et inscrite sur le chemin de ce poste avec sa bibliothèque de types. L'activation COM et un appel `IDispatch` depuis Windows Script Host réussissent. Après un essai manuel qui affichait encore « Impossible de charger le complément CodexVBE », de nouveaux processus Excel lancés nativement ont chargé le complément : le constructeur, `OnConnection` et le tube nommé ont été observés. Sur un nouveau `Classeur1` sans nom, `list_projects` renvoie maintenant `VBAProject` avec `FileName=null` et le mode Conception. Les processus de test ont été fermés proprement. La différence avec l'essai manuel reste inexpliquée ; aucun doublon d'inscription VBE n'a été trouvé dans les vues utilisateur et machine du registre. Dans une ancienne instance SOLIDWORKS, fermer l'éditeur de macro avait seulement masqué sa fenêtre et laissé la DLL verrouillée ; une architecture rechargeable reste à concevoir.

Le 25 septembre 2026, SOLIDWORKS 2019 SP5.0 n'affichait plus CodexVBE dans le Gestionnaire de compléments du VBE malgré une inscription apparemment présente. Une désinstallation complète puis une réinstallation avaient été vérifiées seulement depuis le shell Codex, qui voyait une vue isolée du registre. Un essai avec un nouveau ProgID et un nouveau CLSID avait donc produit le même résultat dans SOLIDWORKS ; l'identité initiale a été restaurée et l'entrée temporaire retirée.

Le 26 septembre, une trace des accès au registre a montré que SOLIDWORKS lisait `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64` et n'y trouvait aucune sous-clé. Les commandes exécutées dans l'environnement Codex voyaient une autre vue du registre (`\REGISTRY\WC\Silo...user_sid`). Après installation dans la vue Windows réelle par une tâche planifiée temporaire, le Gestionnaire de compléments SOLIDWORKS est passé de zéro à une ligne « CodexVBE » ; ses cases « Chargé/déchargé » et « Charger au démarrage » ont été vérifiées cochées. Le nouvel installeur automatise cette voie depuis Codex.

## Suite

Le cycle de point d'arrêt a été vérifié dans Excel avec les identifiants et intitulés natifs relevés dans cette version française du VBE. Le prochain essai pourra porter sur `test.swp` dans SOLIDWORKS une fois son IDE ouvert par l'utilisateur. L'API VBIDE ne fournit pas encore d'inventaire des points d'arrêt ; l'effet de la commande est prouvé par le comportement du débogueur.
