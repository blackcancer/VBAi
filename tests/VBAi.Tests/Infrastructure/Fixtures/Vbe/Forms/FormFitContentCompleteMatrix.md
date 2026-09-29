# Dimensionnement natif — matrice complète avant exécution

- DPI owned réel et frontière bornée : absent/47/48/96/144/768/769, arrondi contenant et conservation du scroll.
- Propriétés Designer/VBIDE : fallback Width/Height, nom/indices/null/type, lecture/écriture, contrats PropertyDescriptor complets, types scalaires numériques et read-only.
- Mesures : null/négatif/non-fini/borne, Controls absent/non énumérable, parent étranger, doublon, 513 enfants.
- Requêtes : padding horizontal/vertical, path/parent/root, mesures internes/externes, dimensions proposées nulles/trop grandes.
- Relecture : largeur/hauteur divergentes, viewport petit/grand sur chaque axe, extent scroll dépassé, drift enfants, setter échoué ; incertitude sans relance.
- Toutes frontières globales sont restaurées en finally ; aucun Office ni réglage système modifié.
