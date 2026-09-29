# Diagnostic comparatif : panneau VBAi invisible après initialisation

## Symptôme à vérifier

Sur le poste concerné, le panneau apparaît brièvement puis disparaît dès la première ouverture du VBE. Le problème est signalé dans Excel et SOLIDWORKS. Ne pas limiter le diagnostic à une fermeture/réouverture.

Le [diagnostic corrigé](../chat-persistence-investigation.md) identifie dans Excel un panneau natif réduit à six pixels par le rattachement automatique au cadre principal. Une correction préserve les dispositions utilisables et récupère les panneaux trop petits dans une fenêtre native flottante ancrable.

## Relevés attendus sur chaque poste

1. Relever l’hôte, sa version, l’architecture, la version de VBE et le SHA-256 de la DLL réellement enregistrée.
2. À la première ouverture, attendre la fin de l’initialisation et capturer le VBE et le panneau. Vérifier l’accès au bas du chat (saisie, fournisseur, modèle), pas seulement sa présence dans `VBE.Windows`.
3. Relever le rectangle natif de `GenericPane` « VBAi » et l’état `VBAi.AddIn.Connect`. Ne pas conclure à partir de `Window.Width/Height` quand il est ancré : ces valeurs peuvent décrire le cadre complet. `Window.HWnd` peut aussi valoir zéro pour le panneau personnalisé.
4. Vérifier une disposition utilisable déjà mémorisée, puis une disposition trop étroite ou écrasée dans une session de test. La première doit être conservée ; la seconde récupérée.
5. Fermer avec la croix du VBE, laisser l’hôte ouvert, puis rouvrir. Ne pas substituer `WM_CLOSE` à la croix : la sonde utilise `WM_SYSCOMMAND / SC_CLOSE`.
6. Relever uniquement les lignes pertinentes de `%TEMP%/VBAi-load.log` (`OnConnection`, récupération du panneau, erreurs d’ancrage, `OnDisconnection`).
7. Fermer normalement la session jetable. Ne pas forcer la reconnexion d’une ancienne référence COM.

## Sonde Excel

Fermer ses sessions Excel avant de lancer :

```powershell
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File .\tools\probes\Test-ChatPanelPlacementExcel.ps1
```

La sonde crée un classeur vide, ne l’enregistre pas et quitte Excel proprement. Elle écrit les états et captures dans `artifacts/chat-panel-placement/`. Le panneau peut changer de position si la disposition enregistrée est trop petite.

## Résultat attendu

Fournir un tableau par hôte et par poste : DLL, première ouverture, dimensions natives, lisibilité du contenu, état connecté, fermeture/réouverture et sortie propre. Masquer chemins personnels, données de macro et identifiants dans les pièces partagées.

**État connu :** récupération et réouverture vérifiées dans Excel sur le poste concerné. SOLIDWORKS et le poste où le panneau fonctionnait auparavant restent à valider.
