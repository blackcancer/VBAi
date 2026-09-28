# Disparition du panneau pendant l’initialisation du VBE

## Périmètre et conclusion — 28 septembre 2026

Le symptôme précisé par l’utilisateur est une disparition **dès la première ouverture, après initialisation**, dans Excel et SOLIDWORKS sur ce poste. La réouverture n’est pas nécessaire pour le provoquer.

Le diagnostic local dans Excel montre un problème de placement du panneau natif : le chat existe, le complément reste connecté, mais son conteneur peut être réduit à une bande de six pixels. Le passage automatique du formulaire flottant au panneau VBE explique le bref affichage initial.

La correction part de `main` au commit `b52e5df`. Les changements de compatibilité Codex/GitHub de la PR #1 sont déjà intégrés à cette base.

## Preuves natives

| Observation dans Excel | Résultat |
|---|---|
| Panneau avant correction | `GenericPane` « VBAi », 1920 × 6 pixels |
| Contrôle enfant | 1 pixel de haut, situé sous la partie visible du panneau |
| État VBIDE | `Visible=true`, malgré le contenu inaccessible |
| `Window.HWnd` du panneau personnalisé | Zéro dans cet hôte ; inutilisable pour conclure à son absence |
| `Window.Width/Height` quand ancré | Dimensions du cadre, pas de la surface réellement réservée au chat |
| Autre disposition mémorisée observée | Largeur 321 pixels, inférieure au minimum du chat (440 × 560) |
| Après récupération | Cadre flottant natif 600 × 820 ; panneau visible 584 × 781 |

Dans `AddIn.ToggleDock`, l’appel systématique à `MainWindow.LinkedWindows.Add(nativeChatWindow)` après `CreateToolWindow` imposait un nouveau rattachement au cadre principal. La disposition mémorisée n’était donc pas simplement restaurée. L’essai natif a montré qu’une affectation de hauteur sur le panneau encore ancré échoue ; après retrait de son `LinkedWindowFrame`, les dimensions flottantes sont applicables.

## Correction

- Laisser `CreateToolWindow` restaurer la disposition associée au GUID stable du panneau.
- Supprimer le rattachement systématique à `MainWindow.LinkedWindows`.
- Mesurer la zone cliente du parent natif de `ChatToolWindow`, avec `GetParent` et `GetClientRect`.
- Conserver la disposition lorsque cette zone respecte le minimum du chat.
- Si elle est trop petite, retirer uniquement ce panneau de son cadre puis lui donner une fenêtre native flottante ancrable de 600 × 820, bornée à la zone de travail de l’écran du VBE.
- Vérifier aussi la taille à la demande d’affichage du chat. Aucun raccourci clavier ni besoin de conserver le focus n’est ajouté.
- En cas d’échec pendant l’attachement initial, conserver le repli existant vers le formulaire flottant avec son message d’erreur.

Le panneau récupéré peut ensuite être ancré manuellement à la position souhaitée. Une disposition utilisable reste inchangée.

## Rectification du diagnostic de fermeture précédent

Les premiers essais utilisaient **`WM_CLOSE`** sur la fenêtre principale. Ils ont effectivement produit `OnDisconnection(0)`, une collection `AddIns` vide et une fenêtre absente après réouverture. Le relevé indépendant de la branche main (Excel PID 33208, 10:04–10:05) employait aussi `WM_CLOSE`.

Ces observations étaient réelles mais **ne reproduisaient pas la croix du VBE dans ce contexte**. La commande système `WM_SYSCOMMAND / SC_CLOSE` masque ici l’éditeur et son panneau tout en conservant le complément connecté. Elles ne justifient donc pas un nouveau mécanisme de reconnexion du complément.

L’ancienne tentative `Connect=true` sur une référence COM périmée a fait planter une instance jetable. Elle n’est ni reprise ni introduite dans la correction.

## Validation

Compilation Debug x64 réussie et **21 tests ciblés réussis sur 21** : attachement, conservation des dimensions utilisables, récupération des dispositions écrasées/trop étroites, garde-fous natifs et repli en cas d’erreur COM. La DLL et la bibliothèque de types ont été installées avec le script officiel, dont la vérification externe a réussi.

Essai réel Excel du 28 septembre 2026, PID 455836 :

| Étape | Heure locale | VBE | Panneau VBAi | Complément |
|---|---|---|---|---|
| Première ouverture | 10:31:31 | Visible | Visible, 584 × 781 | Connecté |
| Croix / SC_CLOSE | 10:31:34 | Masqué | Masqué | Connecté |
| Réouverture | 10:31:37 | Visible | Visible, mêmes dimensions | Connecté |

Le journal confirme la récupération de la zone cliente trop étroite : `321x743 -> 600x820`. La capture montre le chat complet, y compris la saisie et les réglages. Le classeur jetable a été fermé sans sauvegarde et Excel a quitté par `Quit()`, sans arrêt forcé.

Un second lancement complet d’Excel (PID 457240, 10:35:57–10:36:03) a retrouvé la même disposition utilisable dès l’ouverture, puis après SC_CLOSE/réouverture. Cette instance a aussi quitté normalement. Ses preuves sont dans `artifacts/chat-panel-placement-restart/`.

La sonde reproductible est `tools/probes/Test-ChatPanelPlacementExcel.ps1`. Elle refuse une session Excel préexistante, crée son classeur jetable, vérifie présence/dimensions/connexion, capture le rendu, exécute SC_CLOSE puis rouvre le VBE. Elle ferme proprement son instance en `finally`. Elle peut modifier la disposition VBE mémorisée en déclenchant la récupération, mais ne modifie aucune macro.

```powershell
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File .\tools\probes\Test-ChatPanelPlacementExcel.ps1
```

Les relevés JSON et captures sont enregistrés dans `artifacts/chat-panel-placement/` (ignoré par Git). Ces captures doivent être examinées : la seule propriété `Visible` ne prouve pas que le contenu est lisible.

## Limites

- Validation réelle effectuée dans Excel sur ce poste. SOLIDWORKS et un second poste restent à valider ; le code corrigé est commun aux hôtes VBE.
- La mesure s’effectue à l’attachement et à la commande d’affichage. Elle n’impose pas en permanence une taille pendant les redimensionnements manuels.
- Le message de connexion ChatGPT visible sur la capture est distinct du placement ; aucune authentification ni conversation n’a été exécutée par ce test.
- Les anciens relevés basés sur WM_CLOSE restent des observations d’un scénario artificiel, pas une preuve du parcours utilisateur.

## Références

- [CreateToolWindow](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtoolwindow-method) : création de la fenêtre ancrable et identifiant de position.
- [LinkedWindows.Add](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-vba-add-in-object-model) : déplacement entre cadres.
- [WM_SYSCOMMAND](https://learn.microsoft.com/en-us/windows/win32/menurc/wm-syscommand) : commande SC_CLOSE de la barre de titre.
- [WM_CLOSE](https://learn.microsoft.com/en-us/windows/win32/winmsg/wm-close) : demande de fermeture, distincte du scénario testé par la croix.
