# Initialisation et nettoyage des barres / placement du chat

Matrice ajoutée avant exécution :

- AddIn.EnsureUsableChatPlacement : contrôle absent, handle/site absent ou lecture refusée, largeur insuffisante, hauteur insuffisante, deux dimensions exactes suffisantes ; cadre lié absent/présent et refus de détachement ; dispositions natives préservées ou réparées, réouverture réelle.
- AddIn.CleanupTemporaryToolbarCommands : VBE absent, collection normale avec copies temporaire/persistante/étrangère, menus exclus ; suppression native refusée journalisée, arrêt poursuivi ; déconnexion et shutdown effectifs.
- ChatToolWindow.TryGetNativeSiteSize : contrôle sans handle, parent absent, lecture refusée, rectangle normal/dégénéré/inversé, contrôle détruit. Assertions sur size Empty ou dimensions bornées, aucune API d'hôte.
- VbeSession : vrais services créés et vraie restauration SQLite depuis racine temporaire ; processus EXCEL/SLDWORKS sans casse et autre processus, preuve que la racine n'est pas consultée pour les autres hôtes, bases distinctes par hôte et profils existants/vides/corrompus ; erreur de restauration exposée par list_toolbars sans arrêter la session.

La nouvelle surcharge interne injecte uniquement deux lectures système (nom de processus et dossier local) ; les constructeurs publics conservent Process.GetCurrentProcess().ProcessName et Environment.GetFolderPath(LocalApplicationData). Pas de substitution du routage ou de la restauration. Bases SQLite temporaires et doubles COM existants ; globals restaurés par HostUiScope, aucun Excel/SOLIDWORKS réel.

## Preuve ciblée finale

293 tests verts, aucun ignoré ; collector Exclude=[ProviderTests]* uniquement.
Rapport : `artifacts/coverage/host-toolbar-init-final/ec763ce2-1f11-4549-ad1c-9f6447ecffd1/coverage.cobertura.xml`.

| Source | Lignes couvertes | Branches couvertes |
| --- | ---: | ---: |
| Host/AddIn.cs | 191/191 | 158/158 |
| Host/ChatToolWindow.cs | 47/47 | 28/28 |
| Vbe/VbeSession.cs | 522/522 | 1272/1272 |

Test-TestLayout PASS (134 miroirs / 183 sources). Les tests existants de callbacks et de génération ListBinding sont inclus dans le filtre final afin de mesurer les fichiers entiers.
