# Adaptateur EditorVbeModule — matrice avant exécution

Tests miroirs : `Unit/Editor/EditorVbeModule.Tests.cs`. Contrats VBIDE gérés, exports/imports sur vrais fichiers temporaires, aucun Office. La preuve native Excel de la PR9 reste distincte.

- Identité : projets/modules présents ou supprimés, chemins absolus/relatifs/absents/inaccessibles, noms mutables, voisins, composants vides, fenêtre de code, PID sur HWND détenu/absent/invalide.
- Écriture : mode/protection/conflit, plans exacts/périmés, texte non encodable, remplacement/insertion/suppression/no-op, lectures et pannes après mutation.
- Restauration simple : remplacement ou plusieurs lignes, suppression seule/insertion seule, retour divergent, échec de restauration agrégé.
- Attributs simples : déclaration modifiée ou corps voisin, relecture exacte, perte de métadonnées, échec de chargement et récupération, double échec et rétention de l'export.
- Remplacement préparé : standard/classe/formulaire, document exclu, préfixe vide de formulaire, type/code/attributs/concepteur divergeant, mode/texte/nom/concepteur concurrents avant retrait.
- Échec de retrait et de réouverture : candidat présent/absent, import échouant avant/après allocation, original présent/retiré, restauration valide/divergente/échouant, nom concurrent conservé, fichiers de récupération signalés.
- Garde des déclarations : variables, fonctions, propriétés, qualificateurs, déclarations incomplètes, remplacement intersectant/insertion fractionnant/édition adjacente.
