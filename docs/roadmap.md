# Travaux restants

État consolidé au **28 septembre 2026**, avec les extensions fonctionnelles basées sur `b843972f`. Les chiffres de couverture ci-dessous proviennent du bilan `2197c43` et ne mesurent pas les nouvelles classes. Ce document distingue les fonctions présentes des preuves de qualification manquantes. Le [catalogue LLM](reference/vbe-tools.md) décrit l’exposition actuelle ; les [archives](archive/README.md) conservent les expériences détaillées.

## Couverture et documentation

| Travail | État actuel | Critère de fin |
| --- | --- | --- |
| Couverture du code de production | 92,94 % lignes, 87,79 % branches ; 63 classes instrumentées incomplètes après chat-ux | 100 % lignes et branches mesurées sur toute l’assembly, sans masquer du code ; suite globale verte |
| Organisation des tests | 99 miroirs pour 168 sources ; scénarios et fixtures complémentaires | Chaque surface exécutée couverte par son miroir ou un scénario justifié |
| Documentation IntelliSense | Nouvelles déclarations apportées par chat-ux ; agent documentaire en cours | Audit final des déclarations privées/publiques, propriétés et tests, sans lacune |
| Concepteurs WinForms | 24 DesignSurface validées | Préserver cette accessibilité après chaque changement de structure ; qualifier aussi le rendu réel |

Compléter chaque branche de couverture identifiée avant de passer à la suivante. Construire le lot de scénarios cohérent avant de le lancer, puis mesurer la suite globale. Les pourcentages actuels sont détaillés dans [le bilan de tests](test-coverage.md).

## Extensions fonctionnelles du 28 septembre

Les fonctions et les limites exactes sont détaillées dans [Extensions fonctionnelles VBE](reference/functional-extensions.md). Ce lot ajoute 11 outils (177 au total) : exécution paramétrée, renommage local, personnalisation des barres, mutation bornée des options et confiance de certificat. Il étend aussi les déclarations/références, le renommage de projet Excel et la sauvegarde standalone `.swp`. Les tests SOLIDWORKS restent différés.

## Qualification de l’éditeur

| Surface | Ce qui reste à qualifier ou développer |
| --- | --- |
| Modules/classes | Cas d’erreur, encodages/imports et persistance dans les différents hôtes ; ne pas confondre catalogue de commandes et validation de toutes leurs combinaisons |
| Projets | Renommage de projet Excel enregistré/non protégé disponible ; autres périmètres refusés. Sauvegarde standalone `.swp` implémentée mais non qualifiée dans SOLIDWORKS ; autres hôtes non implémentés |
| Signature | Première sélection du certificat dans Sécurité Windows, digest de la signature VBA et persistance SOLIDWORKS ; confiance de certificat disponible hors ligne ; l’état signé ne prouve pas la confiance du certificat |
| Breakpoints | Inventaire indépendant des marqueurs et pointeur d’exécution ; la commande de basculement et l’arrêt effectif sont qualifiés, pas un inventaire exhaustif |
| Variables/espions | Lecture des valeurs SOLIDWORKS, grands arbres, types particuliers et variantes de langues ; dernier essai SOLIDWORKS : zéro ligne exposée |
| Exécution | Appel paramétré implémenté, à qualifier depuis le complément installé ; diagnostics particuliers et variantes d’hôtes ; les statuts asynchrones ne prouvent pas la réussite runtime |
| UserForms | Propriétés réellement modifiables par type, persistance et effet runtime ; contrôles tiers ; fidélité des images, copies et récupérations ; événements et conteneurs complexes |
| Listes | Initialisation multicolonne et liaisons implémentées ; qualifier les combinaisons de contrôles/conteneurs et les autres hôtes |
| Fenêtres/barres d’outils | Persistance, géométries et DPI/écrans ; premier ancrage du chat à droite sur un profil vierge |
| Explorateur d’objets | Lecture/sélection/pagination implémentées ; qualification SOLIDWORKS et variantes UI natives |
| Options/boîte à outils | Mutation bornée Éditeur/Général implémentée, persistance à qualifier ; personnalisation de la boîte à outils absente et contrôles ActiveX tiers à qualifier |

Les journaux anciens peuvent indiquer « manquant » pour des fonctions implémentées depuis : signets, navigation, mise en page, presse-papiers, historique, lancement UserForm, barres d’outils et lecture de l’Explorateur sont désormais dans le code.

## Conversation, fournisseurs et intégration

- Vérifier le rendu et les comportements dans les hôtes réels, au-delà des doubles COM et des captures de démonstration.
- Qualifier les fournisseurs authentifiés et leurs contraintes de modèles/outils séparément des simulations locales ; les crédits ou services locaux ne sont pas supposés disponibles.
- Qualifier interruption/reprise, erreurs de transport et récupération après modification native partielle.
- Poursuivre la relecture linguistique des catalogues ; leur parité technique ne garantit pas la qualité de traduction.

Excel reste l’hôte automatisé prioritaire. Les essais SOLIDWORKS utilisent une instance et un VBE préchargés par l’utilisateur, sans démarrer une nouvelle instance COM. Les essais d’écriture restent limités aux projets jetables identifiés.
