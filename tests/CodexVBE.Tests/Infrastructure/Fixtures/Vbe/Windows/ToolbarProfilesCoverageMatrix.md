# Profils des barres natives — matrice complète

Avant la première exécution :

- VbeToolbarProfiles : collection nulle, bornes 32/33, entrée nulle, noms distincts sans casse ; validation individuelle de chaque terme (nom, caractères, position, coordonnées, rang, commandes, identité, caption, tag), bornes exactes et valeurs optionnelles absentes. Le helper ValidateContents conserve toutes les gardes et accepte une entrée isolée nulle ; aucune garde n'est supprimée.
- ChatSessionStore.ToolbarProfiles : lecture vide et multiple, JSON null/invalide, payload borné, clé normalisée et remplacement, suppression, mismatch, validation originale/nouvelle en échec, payload écrit trop grand, erreurs SQLite BEGIN/INSERT/COMMIT et rollback effectif laissant une transaction réutilisable.
- VbeEditorWindows.ToolbarProfiles : absence de stockage, lecture corrompue, absence/création/réutilisation de barre, ambiguïté/native/menu/protection, chaque identité de contrôle (absent/unique/divergent/dupliqué), sources natives absentes ou incompatibles, copie/tag refusés ; placements flottants complets/partiels et rang de docking présent/absent ; suppression temporaire strictement limitée aux tags GUID minuscules et aux barres normales.
- Sauvegarde via les vraies opérations create/add/remove/show/hide/position/placement : profils existants/absents, temporaires/persistants, géométrie et commandes conservées, refus de persistance pour barre non enregistrée, fautes de stockage signalées après mutation native, caption de copie divergente refusée.
- RemoveToolbar : refus natif conservant le profil, suppression confirmée effaçant seulement le profil ciblé. Correction d'ordre : ne supprimer le profil qu'après preuve d'absence native.

SQLite réel dans un dossier jetable ; doubles COM documentés aux frontières CommandBars/Controls. Aucune fenêtre d'hôte, clipboard, authentification ou réseau. Aucun filtre de production ; collector Exclude=[ProviderTests]* uniquement.

## Résultat ciblé

42 tests verts, aucun ignoré ; Test-TestLayout PASS (133 miroirs / 183 sources).
Rapport `artifacts/coverage/toolbar-profiles-final/9aef6fbb-32b8-48ab-9bbb-82652db410f7/coverage.cobertura.xml`.

| Source | Lignes couvertes | Branches couvertes |
| --- | ---: | ---: |
| ChatSessionStore.ToolbarProfiles.cs | 30/30 | 12/12 |
| VbeToolbarProfiles.cs | 25/25 | 54/54 |
| VbeEditorWindows.ToolbarProfiles.cs | 59/59 | 204/204 |
| VbeEditorWindows.ToolbarCustomization.cs | 113/113 | 284/284 |
| VbeEditorWindows.Toolbars.cs | 148/148 | 390/390 |
