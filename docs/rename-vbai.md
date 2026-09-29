# Renommage VBAi

Le complément, la solution, les projets, l’assembly C#, les namespaces, les ressources, le moteur natif et les outils portent désormais le nom **VBAi**. Le dépôt est `blackcancer/VBAi`. Les identifiants GUID COM restent identiques afin de remplacer la même installation.

Construire `VBAi.sln`, exporter `bin/Debug/net48/VBAi.tlb`, puis utiliser `tools/Install-VBAi.ps1`. Les hôtes VBA doivent être fermés pendant l’installation. Les ProgID deviennent `VBAi.AddIn` et `VBAi.ChatToolWindow` ; le canal local devient `VBAi.<PID>`.

L’installateur reconnaît uniquement l’ancienne identité connue pour migrer son enregistrement. Après installation de la nouvelle identité, il retire l’ancienne entrée du gestionnaire de compléments et les anciens ProgID qui pointent sur les GUID attendus. Les références littérales à l’ancien nom subsistent exclusivement dans cette migration.

Les données des dossiers utilisateur précédents sont copiées vers les dossiers `VBAi` correspondants : configuration, conversations, fournisseurs isolés, brouillons, préférences de thème et autres données persistantes. Un fichier déjà présent dans la destination n’est jamais remplacé ; les sources restent disponibles. Les caches WebView, du moteur natif et Git temporaires ne sont pas copiés. Les liens symboliques sont refusés. Le chiffrement DPAPI reste lié au même utilisateur, sans changement d’entropie.

Le chemin physique de ce checkout reste `E:\Développement\AddIn\CodexVBA`. Les worktrees locaux et leurs travaux ne sont pas supprimés. Les branches distantes autres que `main` sont supprimées à la demande de l’utilisateur ; leurs commits sont conservés localement sous `refs/archive/pre-vbai/`.

Les mesures de couverture antérieures au renommage restent historiques. Le renommage n’est pas une nouvelle mesure de couverture des branches ajoutées par la PR #13.

## Validation du 29 septembre 2026

- Compilation Debug de la solution : aucune erreur ni avertissement, moteur natif reconstruit sous sa nouvelle identité.
- Suite VSTest complète : **1 876 réussites, 0 échec, 21 ignorés**, en **8 min 8 s** (`artifacts/rename-vbai/accepted/global.trx`). Le premier passage est interrompu après une course de démarrage dans le test de streaming ; le passage accepté attend la fin réelle de l’initialisation.
- Les trois scénarios de performance déplacés parmi les scénarios complémentaires passent aussi après reconstruction (`layout-regressions/performance.trx`). La structure comporte **242 miroirs pour 299 sources** (`test-layout.json`).
- Les **46 concepteurs WinForms** sont chargés, redimensionnés et sérialisés (`artifacts/rename-vbai/designers`).
- Installation vérifiée dans la ruche Windows réelle par `Invoke-VBAi-OutsideSandbox.ps1`. Le contrôle d’élévation reconnaît aussi les erreurs d’accès au registre encapsulées par PowerShell. Les premières inscriptions directes dans la vue isolée ne constituent pas une preuve de chargement natif.
- Excel visible charge la DLL attendue, identifiée par son ModuleVersionId, ouvre automatiquement Monaco et conserve le code natif derrière lui. Le redimensionnement et la disposition de l’Explorateur d’objets sont préservés (`artifacts/rename-vbai/excel-startup/monaco-startup.json`). Le classeur jetable et Excel sont fermés ; aucune macro n’est exécutée.
- Le dépôt reste privé et sa seule branche distante est `main` ; **11 branches distantes** sont supprimées. Les probes locaux non suivis restent intacts et ne sont pas publiés.
