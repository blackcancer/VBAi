# Navigation des projets — matrice complète avant exécution

- Découverte : PID, classe racine, titre, arbres imbriqués et doublons ; HWND réellement détenus par le processus test.
- Cycle : actualisation initiale/périodique, fenêtre non projet, doublon, arbre détruit, libération vivante ou déjà détruite, erreur VBIDE.
- Routage : module/classe/document, formulaire et type invalide, sélection absente/inaccessible/invalide, dispatcher fermé, ouverture différée.
- Messages : double-clic synthétique sur HWND test, région autorisée/refusée, élément nul, sélection divergente, routage accepté/refusé. Aucune interaction souris sur un hôte réel.
- Appartenance : arbre créé sur un autre STA du processus trouvé mais jamais sous-classé par le thread courant.
