# Travaux restants

État consolidé le **28 septembre 2026**, après la PR #6 et le lot de complétion IDE. Les anciens passages à 100 % ne constituent pas la mesure du code actuel : voir [le bilan de tests](test-coverage.md). Ce document distingue les contrats présents de la qualification native manquante. Le [catalogue LLM](reference/vbe-tools.md) expose 197 outils ; les [archives](archive/README.md) conservent les expériences détaillées.

## Couverture et documentation

| Travail | État actuel | Critère de fin |
| --- | --- | --- |
| Couverture du code de production | Suite globale verte ; 97,35 % lignes et 95,94 % branches ; 648 lignes et 1 045 branches restantes | 100 % lignes et branches mesurées sur toute l’assembly, sans masquer du code ; suite globale verte |
| Organisation des tests | 171 miroirs pour 228 sources ; scénarios et fixtures complémentaires | Chaque surface exécutée couverte par son miroir ou un scénario justifié |
| Documentation IntelliSense | Travail Luna conservé séparément, intégration et conflits encore à traiter | Audit final des déclarations privées/publiques, propriétés et tests, sans lacune |
| Concepteurs WinForms | 31 DesignSurface validées | Préserver cette accessibilité après chaque changement de structure ; qualifier aussi le rendu réel |

Compléter chaque branche de couverture identifiée avant de passer à la suivante. Construire le lot de scénarios cohérent avant de le lancer, puis mesurer la suite globale. Les pourcentages actuels sont détaillés dans [le bilan de tests](test-coverage.md).

## Extensions fonctionnelles du 28 septembre

Les fonctions et les limites exactes sont détaillées dans [Extensions fonctionnelles VBE](reference/functional-extensions.md). Le lot initial ajoutait 11 outils ; le catalogue actuel en contient 197. Contrats historiques : exécution paramétrée, renommage local, personnalisation des barres, mutation bornée des options et confiance de certificat. Il étend aussi les déclarations/références, le renommage de projet Excel et la sauvegarde standalone `.swp`. Les tests SOLIDWORKS restent différés.

## Qualification de l’éditeur

L'[inventaire fonctionnel complet](reference/vbe-capability-inventory.md) couvre désormais toutes les surfaces de l'éditeur. Le lot IDE `52cb537` ajoute trois contrats (180 outils au total), le renommage borné de paramètres privés et les arguments nommés de `run_procedure`. La vérification du fichier signé est implémentée mais son SIP natif attend une autorisation distincte. La PR #4 ajoute les structures fixes Designer ; elle est incluse dans la mesure globale `27389a8`, revenue à 100 % lignes et branches.

| Surface | Ce qui reste à qualifier ou développer |
| --- | --- |
| Modules/classes | Cas d’erreur, encodages/imports et persistance dans les différents hôtes ; ne pas confondre catalogue de commandes et validation de toutes leurs combinaisons |
| Projets | Renommage de projet Excel enregistré/non protégé disponible ; autres périmètres refusés. Sauvegarde standalone `.swp` implémentée mais non qualifiée dans SOLIDWORKS ; adaptateurs Word/PowerPoint implémentés mais NOT_RUN ; cycle Add/Open/Remove standalone et protection ajoutés, protection sauvegardée/réouverte qualifiée dans Excel |
| Signature | Première sélection du certificat dans Sécurité Windows, digest de la signature VBA et persistance SOLIDWORKS ; confiance de certificat disponible hors ligne ; l’état signé ne prouve pas la confiance du certificat |
| Breakpoints | Inventaire indépendant des marqueurs et pointeur d’exécution ; la commande de basculement et l’arrêt effectif sont qualifiés, pas un inventaire exhaustif |
| Variables/espions | Lecture des valeurs SOLIDWORKS, arbres COM/espions et variantes de langues ; tableau Excel de 1 000 éléments et sept types scalaires qualifiés en français, doublons UIA supprimés ; dernier essai SOLIDWORKS : zéro ligne exposée |
| Exécution | Scalaires, appels nommés et transport Excel.Run à invocation unique pour tableaux/retours scalaires qualifiés dans Excel ; classes/objets COM/tableaux ByRef non couverts ; autres hôtes à qualifier |
| UserForms | Propriétés réellement modifiables par type, persistance et effet runtime ; contrôles tiers ; fidélité des images, copies et récupérations ; événements et conteneurs complexes |
| Listes | Initialisation multicolonne et liaisons implémentées ; qualifier les combinaisons de contrôles/conteneurs et les autres hôtes |
| Fenêtres/barres d’outils | Persistance des barres Excel qualifiée via SQLite ; géométries/DPI/ancrage natifs au-delà du profil mesuré ; premier ancrage du chat à droite sur un profil vierge |
| Explorateur d’objets | Lecture/sélection/pagination implémentées ; qualification SOLIDWORKS et variantes UI natives |
| Options/boîte à outils | Éditeur/Général et cases Format/Ancrage qualifiés après réouverture et restauration ; choix de police/palettes et autres langues natives à qualifier. Personnalisation de la boîte à outils absente ; neuf MSComctl installés refusés par la politique native de confiance |

Les journaux anciens peuvent indiquer « manquant » pour des fonctions implémentées depuis : signets, navigation, mise en page, presse-papiers, historique, lancement UserForm, barres d’outils et lecture de l’Explorateur sont désormais dans le code.

## Conversation, fournisseurs et intégration

- Vérifier le rendu et les comportements dans les hôtes réels, au-delà des doubles COM et des captures de démonstration.
- Qualifier les fournisseurs authentifiés et leurs contraintes de modèles/outils séparément des simulations locales ; les crédits ou services locaux ne sont pas supposés disponibles.
- Qualifier interruption/reprise, erreurs de transport et récupération après modification native partielle.
- Poursuivre la relecture linguistique des catalogues ; leur parité technique ne garantit pas la qualité de traduction.

Excel reste l’hôte automatisé prioritaire. Les essais SOLIDWORKS utilisent une instance et un VBE préchargés par l’utilisateur, sans démarrer une nouvelle instance COM. Les essais d’écriture restent limités aux projets jetables identifiés.

Voir le [bilan précis des qualifications natives](reference/native-qualification.md) pour les scénarios, preuves et conditions restantes.
