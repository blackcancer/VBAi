# Couverture ChatWindow

> Archive conservée le 28 septembre 2026. Ce document contient des observations et des décisions de sa période de rédaction ; ses états « à faire » et ses anciens chiffres ne constituent pas le bilan actuel. Voir [la documentation actuelle](../../README.md) et [les travaux restants](../../roadmap.md).

Le lot global du 27 septembre 2026 a été exécuté : 219 tests réussis, 2 ignorés. Son rapport est `artifacts/coverage/batch-comprehensive-final/9bf3adb1-a9d7-4eac-bc7c-3913122c6063/coverage.cobertura.xml`. Les scénarios ci-dessous font partie de cette campagne ; les limites indiquées restent à couvrir.

| Fichier | Scénarios apportés par `ChatWindowStateTests` | Branches dépendantes d'un environnement actif |
| --- | --- | --- |
| `ChatWindow.cs` | Construction du formulaire seul, choix d'effort selon provider/modèle, état occupé et statut | Constructeur avec session, chargement/sauvegarde des paramètres, catalogue de modèles Codex et autres fournisseurs, envoi/arrêt d'une requête, dialogues Settings et gestion réelle du client. |
| `ChatWindow.Designer.cs` | Création et destruction des contrôles WinForms avec surfaces WPF sous STA | Disposition visuelle réelle, ancrage et événements de la fenêtre hébergée dans le VBE. |
| `ChatWindow.Shell.cs` | Initialisation des modes, changement de mode, préparation d'une commande d'éditeur et garde occupée | Menus natifs, fenêtre GitHub, compilation et raccourcis clavier avec un VBE lié. |
| `ChatWindow.Sessions.cs` | Filtrage et ordre de l'historique, renommage, brouillon sauvegardé sans store, réparation d'appels d'outil interrompus | SQLite utilisateur, projets réels, changement de portée, reprise de thread distant, archivage et mémoire persistée. |
| `ChatWindow.Transcript.cs` | Flux final mis à jour, suppression du doublon final, nettoyage, rendu Markdown et vue des entrées | Presse-papiers, navigation VBIDE, retour arrière des modifications, suivi du défilement après mesure visuelle. |
| `ChatWindow.Composer.cs` | Frontières de jeton, références exactes et dédoublonnées, puces de mémoire et pièce jointe, limite de contexte | Résolution de références dans un vrai projet et sélection clavier/souris de la popup. |
| `ChatWindow.Workflow.cs` | Pièces jointes locales, limite de 48 000 caractères, aperçu du contexte et préparation d'une commande | Sélection active du code VBE, compilation native, export avec dialogue, restauration de code et branche de conversation liée au client. |

Les tests créent `ChatWindow()` puis initialisent séparément shell, composer et transcript. Ils ne construisent pas `ChatWindow(VbeSession)` : celui-ci ouvre l'historique sous AppData et démarre le chargement asynchrone des modèles. Un test fiable de ce constructeur et des branches réseau demande une injection du store, du catalogue et du client ; aucune requête externe ni donnée utilisateur ne doit être touchée par la catégorie `Unit`.
