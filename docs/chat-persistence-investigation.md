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
