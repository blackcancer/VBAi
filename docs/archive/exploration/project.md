# CodexVBE — objectif et état du projet

> Archive conservée le 28 septembre 2026. Ce document contient des observations et des décisions de sa période de rédaction ; ses états « à faire » et ses anciens chiffres ne constituent pas le bilan actuel. Voir [la documentation actuelle](../../README.md) et [les travaux restants](../../roadmap.md).

_État au 25 septembre 2026. Projet expérimental pour le VBE 64 bits, notamment dans Excel et SOLIDWORKS, avec .NET Framework 4.8._

## Objectif

Permettre à Codex de travailler directement dans le **VBE**, quel que soit son hôte compatible, notamment Excel ou SOLIDWORKS, sans cycle systématique d'exportation des modules et de réimportation. L'agent doit pouvoir lire et modifier le projet ouvert, lancer le code par les commandes internes du VBE, observer le résultat, diagnostiquer les erreurs et poursuivre le débogage. Les essais autonomes se font d'abord dans un classeur Excel jetable. La macro `test.swp` reste réservée aux essais SOLIDWORKS lorsque l'utilisateur a ouvert son IDE.

Le système visé comprend :

1. **Exploration** : projets ouverts, modules, procédures, classes, définitions, références et recherche textuelle.
2. **Édition** : lecture et modification ciblée du code vivant, contrôle de concurrence avant écriture, comparaison et retour arrière des changements de l'agent, enregistrement explicite de la macro.
3. **Exécution** : lancement d'une procédure depuis le VBE, sans `ISldWorks.RunMacro2` pour les tests de débogage ; accès à la fenêtre Exécution pour évaluer une expression ou `Debug.Print`, puis lecture fiable de sa sortie.
4. **Débogage** : pose, retrait et localisation des points d'arrêt ; lancement, poursuite, arrêt et pas à pas ; lecture du mode du projet, de la ligne courante, des variables locales et des espions.
5. **Intégration Codex** : commandes structurées utilisables sans focus clavier ni manipulation manuelle répétée du VBE, avec erreurs exploitables et garde-fous sur le projet et la ligne ciblés.

Ces points décrivent **la cible**, pas les capacités déjà terminées. Le VBE ne fournit pas une API publique complète pour toutes ces fonctions ; la lecture des variables locales et de la sortie de la fenêtre Exécution est vérifiée dans Excel, tandis que l'évaluation d'expressions et les espions demandent encore des essais.

## Contraintes décidées

- Les tests d'exécution doivent utiliser le moteur et les commandes internes du VBE. Les API d'exécution de l'hôte, dont `ISldWorks.RunMacro2`, ne répondent pas au besoin de débogage interactif.
- Les raccourcis clavier, `SendKeys`, `Ctrl+G`, `F9` et toute solution qui exige que le VBE garde le focus sont proscrits, y compris pendant les essais.
- Excel est l'hôte des essais autonomes : Codex peut le démarrer et ouvrir son VBE. Pour SOLIDWORKS, l'utilisateur ouvre lui-même l'IDE avant les essais.
- `test.swp` est jetable. `test-backups/test.swp` est une copie de secours ; ne pas la restaurer tant que la macro ouverte reste accessible, afin d'éviter des réouvertures manuelles à chaque essai.
- Ne pas intervenir sur les autres macros ouvertes ou les fichiers de production. La correspondance entre nom VBA et chemin du `.swp` doit être contrôlée avant les écritures et l'exécution.
- L'architecture doit rester exploitable depuis Visual Studio Community via `CodexVBE.sln` en x64.

## Architecture actuelle

| Élément | Rôle |
| --- | --- |
| `CodexVBE.sln`, `src/CodexVBE/CodexVBE.csproj` | Solution Visual Studio Community, cible `net48`, plateforme x64. |
| `src/CodexVBE/Host/AddIn.cs` | Point d'entrée COM du complément VBE ; démarre la passerelle dans le processus hôte du VBE. |
| `src/CodexVBE/Bridge/BridgeServer.cs` | Tube nommé `CodexVBE.<PID hôte>` limité au compte Windows courant ; transfert de requêtes JSON vers le thread du VBE. |
| `src/CodexVBE/Vbe/VbeSession.cs` | Inventaire des projets/modules, lecture, remplacement de lignes et distribution des commandes. |
| `src/CodexVBE/Vbe/Debug/VbeDebug.cs` | Sélection d'une ligne par `CodePane.SetSelection`, inventaire des contrôles VBE et invocation native par `CommandBarControl.Execute`. |
| `src/CodexVBE/Bridge/Protocol.cs` | Contrat JSON des requêtes et réponses. |
| `tools/Invoke-CodexVBE.ps1` | Client de la passerelle intégrée. |
| `tools/VbeController/Invoke-VbeController.ps1` | Contrôleur externe en lecture seule pour l'état, les fenêtres, les recherches, les variables locales et la sortie de la fenêtre Exécution. |
| `tools/Install-CodexVBE.ps1`, `tools/Test-CodexVBEInstallation.ps1` | Inscription COM utilisateur courant et diagnostic. |

Les commandes de la passerelle sont `status`, `list_projects`, `list_modules`, `read_module`, `replace_lines`, `debug_state`, `list_commands`, `select_code` et `invoke_debug`. `replace_lines` exige une empreinte SHA-256 du module et le mode Conception. `invoke_debug` contrôle l'empreinte, la ligne, le mode attendu, le volet actif et l'identité du contrôle VBE. Les actions prévues à ce stade sont `toggle_breakpoint`, `run`, `continue`, `step_into` et `step_over` ; elles ne sont pas toutes testées.

Le contrôleur externe fournit `state`, `windows`, `find_text`, `list_symbols`, `locals` et `immediate`. Il n'envoie aucun raccourci clavier.

## Résultats vérifiés

| Domaine | Résultat observé | Limite |
| --- | --- | --- |
| Installation | Inscription COM et activation de `CodexVBE.AddIn` validées ; la nouvelle DLL contenant `VbeDebug` a été copiée sous `bin/Debug/net48/` après fermeture de SOLIDWORKS. | La nouvelle version n'a pas encore été chargée et exercée dans SOLIDWORKS. |
| Passerelle SOLIDWORKS | Le complément a exposé son tube nommé ; `status`, projets, modules, lecture et modification ciblée ont répondu sur un `.swp` ouvert. | Ces essais portaient sur la version précédente de la DLL. |
| Édition | Une retouche temporaire de `test1` dans `test.swp` a été relue puis annulée dans le VBE ; l'empreinte finale du code vivant égalait l'initiale. | Aucun nouvel enregistrement du `.swp` sur disque n'a été prouvé. |
| Recherche | `list_symbols` a trouvé `test1.main` à la ligne 2 du projet VBA `test`. | Recherche avancée par classe, définition ou références encore à concevoir. |
| Fenêtre Exécution | Dans Excel, `immediate` a lu `2`, produit par `Debug.Print`, avec UI Automation `TextPattern`, sans focus clavier ni coordonnées. | L'évaluation d'une expression saisie dans cette fenêtre et la validation dans SOLIDWORKS restent à faire. |
| Variables locales | Dans Excel, `locals` a lu les lignes de la fenêtre Variables locales avec UI Automation `ValuePattern` ; `probeValue` valait `1` à l'arrêt sur la ligne 4 et `2` après un pas détaillé. | Les valeurs sont actuellement renvoyées comme texte brut des lignes affichées ; les espions et SOLIDWORKS ne sont pas validés. |
| Débogage Excel | Dans un `Classeur1` jetable, la commande native « Basculer le point d'arrêt » (ID 51) a ciblé la ligne 4. « Exécuter Sub/UserForm » (ID 186) a placé le projet en mode Arrêt (`1`) avec la sélection sur la ligne 4. « Pas à pas détaillé » (ID 188) a déplacé la sélection à la ligne 5. « Continuer » (ID 186) a ramené le projet en mode Conception (`2`). Le point d'arrêt a été retiré, puis une nouvelle exécution s'est terminée en mode Conception sans nouvel arrêt. | L'API VBIDE ne fournit pas d'inventaire direct des points d'arrêt. Le changement de mode et de ligne, puis la réexécution sans arrêt, établissent l'effet de cette séquence dans Excel. |
| Fermeture Excel | Après correction du script d'essai, les essais « inspection » puis « pose/retrait » ont terminé ; aucun processus Excel ni nouvel événement de plantage n'a été observé après chacun. | La cause exacte des plantages antérieurs n'est pas établie. La correction évite la libération manuelle des objets COM partagés, masque le VBE, ferme le classeur puis appelle `Quit()`. |
| Chargement sur ce poste | La DLL recompilée et `CodexVBE.tlb` sont inscrites pour l'utilisateur courant. L'activation COM et un appel `IDispatch` réussissent en 64 bits. Plusieurs processus Excel lancés nativement ont chargé CodexVBE dans VBE ; sur un `Classeur1` sans nom, `list_projects` renvoie `VBAProject`, `FileName=null`, `Mode=2`. | Un essai manuel après l'inscription de la bibliothèque de types a encore affiché l'erreur. Aucun double enregistrement VBE n'a été trouvé dans les vues du registre vérifiées. Les essais SOLIDWORKS actuels sont détaillés dans la ligne suivante. |
| Réinstallation et découverte dans SOLIDWORKS | La trace a révélé que les réinstallations depuis Codex modifiaient une vue isolée du registre. Après installation dans la ruche Windows réelle, le Gestionnaire de compléments SOLIDWORKS a affiché « CodexVBE » et ses cases « Chargé/déchargé » et « Charger au démarrage » étaient cochées. | Les commandes de la nouvelle version du complément n'ont pas encore été exercées dans SOLIDWORKS. |

Le Gestionnaire de compléments du VBE SOLIDWORKS a été identifié par sa fenêtre native et sa liste `SysListView32` contenait **zéro ligne**. La session SOLIDWORKS tourne en 64 bits sous le même compte que l'inscription utilisateur ; sa version de `VBE7.DLL` est `7.01.1135`, contre `7.01.1158` dans Office. Deux traces noyau puis Process Monitor ont montré que SOLIDWORKS ouvre `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64` et que `RegEnumKey` sur l'index 0 renvoie `STATUS_NO_MORE_ENTRIES` (0x8000001A). L'inscription visible depuis les commandes Codex se trouvait dans une vue isolée du registre : Process Monitor a identifié pour `reg.exe` le chemin `\REGISTRY\WC\Silo...user_sid\Software\Microsoft\VBA\VBE\6.0\Addins64`. Une tâche planifiée temporaire, lancée hors de cette isolation, a confirmé que l'emplacement utilisateur réel était vide, puis y a exécuté `Install-CodexVBE.ps1`. Après réouverture du Gestionnaire de compléments par la commande accessible native, sa liste contenait une ligne « CodexVBE » ; après activation par l'utilisateur, « Chargé/déchargé » et « Charger au démarrage » étaient cochés. La commande `status` du tube du processus SOLIDWORKS (PID 40840) a ensuite répondu `Connected=true`, version `0.1.0`. Le script d'installation redirige maintenant les appels du shell Codex vers une tâche temporaire et vérifie indépendamment l'inscription ; un appel complet de l'entrée standard puis du diagnostic a réussi.

Les anciens essais Excel provoquaient des plantages `EXCEL.EXE` signalés dans `combase.dll` (`0xc0000005`) lors de la séquence de fermeture. Cette observation ne démontre pas quel appel précis causait le défaut. Le script de travail corrigé se trouve dans le dossier temporaire de la tâche Codex et ne fait pas partie du projet livré.

## Étape en cours

**Étendre la lecture du débogueur, sans focus clavier.** Le chargement, un cycle complet de point d'arrêt, la lecture des variables locales et celle de la sortie de `Debug.Print` sont validés dans Excel. Il reste à examiner les espions et l'évaluation d'expressions dans la fenêtre Exécution. Les classeurs de débogage temporaires ont été fermés sans enregistrement.

Ensuite, la même chaîne pourra être exécutée dans le VBE SOLIDWORKS sur `test.swp`, une fois son IDE ouvert par l'utilisateur. Les intitulés et identifiants des contrôles seront relevés dans **ce** VBE avant invocation ; les valeurs vues dans Excel ne sont pas présumées identiques. Aucune validation SOLIDWORKS de la nouvelle DLL n'est revendiquée.

## Étapes suivantes proposées

1. **Validation du débogage natif** : arrêt effectif, pas à pas, poursuite, retrait du point d'arrêt et gestion des modes ; vérification sur `test.swp`.
2. **Observabilité** : compléter la lecture des variables, examiner les espions et l'évaluation d'expressions dans la fenêtre Exécution ; documenter explicitement ce qui est inaccessible par les interfaces publiques du VBE.
3. **Navigation et édition robuste** : enrichir la recherche de symboles et les opérations de modification, avec vérification du fichier cible, journal des changements et sauvegarde explicite.
4. **Interface agent** : exposer les opérations comme outils Codex contrôlés, avec réponses structurées, temporisations et diagnostic des commandes bloquées.
5. **Distribution** : rendre le chargement et la mise à jour du complément reproductibles, réduire le besoin de redémarrer SOLIDWORKS, préparer les essais sur les configurations cibles.

Le cycle Excel de l'étape 1 est validé ; sa répétition dans SOLIDWORKS exige que l'utilisateur ouvre son IDE. Les prochaines découvertes doivent distinguer l'appel COM réussi, l'effet visible dans le débogueur et le comportement du code réellement exécuté.
