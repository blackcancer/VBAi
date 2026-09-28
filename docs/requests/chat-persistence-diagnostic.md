# Demande de diagnostic — panneau VBAi absent après réouverture du VBE

## Objet

Comparer un poste où le panneau VBAi persiste à celui où il disparaît après fermeture puis réouverture de l’éditeur VBA, sans fermer l’application hôte. Le problème est observé sur le poste concerné dans Excel et SOLIDWORKS. Cette demande vise à identifier une différence de chargement ou de configuration avant de modifier le cycle de vie du complément.

## Constat déjà établi sur le poste concerné

Dans une instance Excel jetable, `VBE.Windows` contient une fenêtre `VBAi` visible à la première ouverture. La fermeture native du VBE déclenche `OnDisconnection(0)` et supprime cette fenêtre. À la réouverture dans le même processus Excel, aucun `OnConnection` n’est journalisé et `VBE.AddIns.Count` vaut zéro. `VBE.AddIns.Update()` ne change pas ce résultat. L’inscription utilisateur `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\CodexVBE.AddIn` reste présente avec `LoadBehavior=3`.

Une tentative de remettre `Connect=true` sur une référence COM conservée avant la fermeture a bloqué puis fait planter l’instance Excel jetable. Ne pas reproduire cette manipulation.

## Relevés demandés sur les deux postes

Effectuer les relevés sous le même compte Windows que celui qui utilise le complément. Dans Excel, utiliser un classeur jetable ; dans SOLIDWORKS, ne modifier aucune macro de production. Ne pas réinstaller le complément, modifier le registre, forcer `Connect` ou fermer une session de travail pour ce diagnostic.

1. Noter les versions d’Excel ou de SOLIDWORKS, de `VBE7.DLL` chargé par le processus et de `CodexVBE.dll` installé.
2. Noter le PID de l’hôte. Ouvrir son VBE une première fois et relever : présence et visibilité de `VBAi`, nombre de fenêtres VBE, nombre de compléments dans `VBE.AddIns`, état `Connect` de `CodexVBE.AddIn` s’il figure dans la collection.
3. Fermer uniquement le VBE par sa croix. Garder Excel ou SOLIDWORKS ouvert. Relever l’heure exacte et les nouvelles lignes `OnDisconnection` du journal `%TEMP%\CodexVBE-load.log`.
4. Rouvrir le VBE depuis l’hôte, sans redémarrer l’hôte. Relever les mêmes quatre états et les éventuelles lignes `OnConnection` du journal. Noter si **Affichage → Assistant VBAi** est présent.
5. Lire les valeurs `FriendlyName`, `LoadBehavior` et la présence de la clé `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\CodexVBE.AddIn`. Indiquer séparément si une inscription existe aussi dans la vue machine 64 bits. Ne pas exporter de clés de registre complètes.
6. Indiquer si la fermeture du VBE décharge `CodexVBE.dll` du processus hôte, si cela peut être observé sans outil intrusif.

Pour Excel, les états du VBE peuvent être lus dans un PowerShell 64 bits, attaché à une session Excel **déjà ouverte** :

```powershell
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $excel.VBE
$chat = @($vbe.Windows | Where-Object { $_.Caption -eq 'VBAi' } | ForEach-Object { [pscustomobject]@{ Caption = $_.Caption; Visible = $_.Visible } })
$addins = @($vbe.AddIns | ForEach-Object { [pscustomobject]@{ ProgId = $_.ProgId; Connect = $_.Connect } })
[pscustomobject]@{ VbeVisible = $vbe.MainWindow.Visible; WindowCount = $vbe.Windows.Count; Chat = $chat; AddIns = $addins } | Format-List
```

Exécuter le bloc une fois à la première ouverture, puis une fois après réouverture du VBE. Si plusieurs instances Excel tournent, ne pas utiliser `GetActiveObject` pour conclure sur une instance précise ; noter cette limite et relever les états dans une session isolée.

## Résultat attendu du diagnostic

Fournir un tableau « poste concerné / poste fonctionnel » avec les états avant fermeture et après réouverture, les versions, les seules lignes pertinentes du journal et la différence constatée. Masquer les noms d’utilisateur, chemins personnels, identifiants de compte et données de macro. Conclure seulement sur les différences observées ; proposer ensuite le plus petit essai de correction réversible et son retour arrière.

Voir aussi [l’investigation sur ce poste](../chat-persistence-investigation.md).
