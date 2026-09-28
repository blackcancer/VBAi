# Persistance du panneau de discussion après fermeture du VBE

## Objectif

Conserver le panneau VBAi lorsque l’éditeur VBA est fermé puis rouvert dans le même processus hôte. Le problème est signalé sur un poste, dans Excel et SOLIDWORKS ; selon l’utilisateur, le panneau persiste sur d’autres postes.

## Reproduction sur le poste concerné

Le 28 septembre 2026, une instance Excel isolée a été lancée avec un classeur jetable. Le VBE a chargé `CodexVBE.AddIn` et `VBE.Windows` contenait la fenêtre visible `VBAi`. La fermeture native de la fenêtre principale du VBE a supprimé cette fenêtre. Après réouverture par `Excel.Application.VBE.MainWindow.Visible = true`, `VBAi` n’était plus présent.

Le journal indique `OnDisconnection(0)` lors de la fermeture du VBE, sans nouvel `OnConnection` lors de sa réouverture dans le même processus Excel. Après réouverture, `VBE.AddIns.Count` vaut zéro. `VBE.AddIns.Update()` ne repeuple pas la collection. L’inscription reste présente sous `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\CodexVBE.AddIn`, avec `LoadBehavior=3`.

L’utilisateur a reproduit la disparition du panneau dans SOLIDWORKS sur le même poste. SOLIDWORKS n’a pas été automatisé dans cette investigation.

## Essais et limites

Masquer puis réafficher le VBE avec `MainWindow.Visible` conserve le panneau : la transition pertinente est la fermeture native de l’éditeur. Une surveillance de visibilité dans le complément n’a pas restauré la fenêtre, car celui-ci reçoit `OnDisconnection` à la fermeture. Cette tentative a été retirée du code, puis la version stable a été recompilée et réinstallée.

Une tentative de passer à `Connect=true` sur une ancienne référence COM après réouverture a bloqué puis fait planter l’instance Excel jetable. Cette méthode de reconnexion ne doit pas être utilisée sur une session de travail. Aucun classeur ni fichier macro n’a été modifié pendant l’essai.

Ces observations établissent le comportement local, mais n’expliquent pas encore sa différence avec les autres postes. Une correction future nécessitera un déclencheur de cycle de vie porté par l’application hôte, ou l’identification d’une différence de configuration. Elle devra être validée séparément dans des sessions jetables Excel et SOLIDWORKS avant déploiement.

## État actuel

La surveillance expérimentale ne figure pas dans la version installée. Les tests ciblés du complément passent (14/14), la bibliothèque de types COM a été régénérée et les scripts officiels d’installation et de vérification ont réussi. La persistance du panneau après fermeture native du VBE reste non résolue sur ce poste.

## Relevé indépendant après intégration de la PR #1

Le 28 septembre 2026, le diagnostic a été reproduit dans une seconde instance Excel jetable, PID `33208`. Aucun fichier de production ni valeur du registre n'a été modifié. Fermeture native par `WM_CLOSE`, puis réouverture par `CommandBars.ExecuteMso("VisualBasic")`, sans raccourci ni coordonnées souris.

| Relevé | Poste observé, première ouverture | Poste observé, réouverture | Poste fonctionnel |
|---|---|---|---|
| Heure locale | 10:04:50.697 +02:00 | 10:05:05.428 +02:00 | Non relevé |
| VBE visible | Oui | Oui | Non relevé |
| Fenêtres VBE | 8 | 7 | Non relevé |
| Fenêtre VBAi | Présente, visible | Absente | Non relevé |
| Compléments `VBE.AddIns` | 1 | 0 | Non relevé |
| `CodexVBE.AddIn.Connect` | `true` | Entrée absente | Non relevé |
| Menu Assistant VBAi | Présent | Absent | Non relevé |

Versions observées : Excel `16.0.20326.20158`, VBE7 `7.01.1135`, assembly et fichier CodexVBE `0.1.0.0`. Le SHA-256 de la DLL enregistrée est `53485764918F4507317D6E56E94422A489D947E6470630450E5E3DD546EF4528`, identique au fichier de sortie Debug présent au début du diagnostic. Cet essai porte sur cette DLL enregistrée ; le build isolé du merge de la PR n'a pas remplacé l'installation.

L'inscription utilisateur `VBA\VBE\6.0\Addins64\CodexVBE.AddIn` est toujours présente : `FriendlyName=CodexVBE`, `LoadBehavior=3`. Elle est visible dans les lectures Registry64 et Registry32 ; cela ne prouve pas deux inscriptions distinctes. Aucune entrée correspondante dans `Addins`, ni inscription machine dans les emplacements inspectés.

Lignes pertinentes du journal :

```text
2026-09-28T10:04:35.3265742+02:00 Constructed: EXCEL PID=33208
2026-09-28T10:04:35.3405742+02:00 OnConnection: EXCEL PID=33208
2026-09-28T10:04:35.3975752+02:00 Bridge started: CodexVBE.33208
2026-09-28T10:04:50.8877951+02:00 OnDisconnection: 0
```

Aucun nouvel `OnConnection` pour ce PID dans l'intervalle de réouverture. Les tests unitaires écrivent aussi dans ce journal ; leurs autres événements ne sont pas attribués à Excel. L'énumération des modules natifs ne permet pas de conclure au déchargement de l'assembly managée CodexVBE.

Les relevés avant/après sont sauvegardés sous `artifacts/panel-lifecycle-diagnostic/`. La finalisation du script est restée bloquée après les relevés ; seule l'instance Excel créée pour ce diagnostic et son processus PowerShell dédié ont été arrêtés. Aucun autre processus hôte n'a été fermé. SOLIDWORKS n'a pas été testé.

### Conclusion et prochain essai réversible

La perte du panneau est reproduite indépendamment. Le complément est déconnecté à la fermeture du VBE, puis le VBE rouvert n'énumère aucun complément alors que l'inscription existe encore. La différence entre postes et la cause de cette collection vide restent inconnues : aucun relevé d'un poste fonctionnel n'est disponible.

Le plus petit essai comparatif suivant consiste à exécuter le même relevé sur le poste fonctionnel, en vérifiant d'abord une DLL identique par SHA-256. Si les DLL diffèrent, tester la même version dans une instance jetable sur les deux postes avant de changer le cycle de vie. Retour arrière : quitter l'instance jetable et rétablir la version initiale avec son installateur si une version d'essai a été installée. Aucune reconnexion via une ancienne référence COM, aucune modification du registre et aucune surveillance expérimentale n'ont été introduites ici.
