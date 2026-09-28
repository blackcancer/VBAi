# Inventaire fonctionnel du VBE

Relecture du code et du catalogue LLM le 28 septembre 2026. Cet inventaire concerne l'éditeur entier, pas seulement le concepteur et le débogueur. Les outils Git et les fournisseurs LLM sont des extensions du complément. La couverture instrumentée du code ne mesure pas cet inventaire fonctionnel.

**Disponible** : contrat présent. **Partiel** : périmètre explicitement limité. **Absent** : fonction spécifique à développer. **À qualifier** : implémentation présente, preuve native insuffisante. Excel et SOLIDWORKS ont des résultats de qualification distincts.

## Surfaces de l'éditeur

| Surface | Contrats disponibles | Manque ou limite restante |
| --- | --- | --- |
| Fichier / projets | Liste des projets ; propriétés ; sauvegarde du document Excel et des projets standalone `.swp` | Ouverture/création/fermeture de projets standalone, impression et adaptateurs de sauvegarde des autres hôtes absents ; rechargement `.swp` à qualifier dans SOLIDWORKS |
| Fichier / composants | Création de module, classe, UserForm ; renommage, suppression, import/export `.bas`, `.cls`, `.frm` et ressources associées ; inspection/insertion de fichier | Formats/encodages et ressources à qualifier sur les autres hôtes ; export d'un composant ne prouve pas la persistance du projet hôte |
| Édition / source | Lecture SHA, remplacement de lignes/procédures, éditions avec prévisualisation, rechercher, presse-papiers, commentaires/indentation, historique géré | Complétion native et informations de paramètres ne constituent pas des contrats spécifiques ; analyse sémantique de projet complète absente |
| Édition / renommage | Variable/constante locale explicite ; paramètre de Sub/Function privée standard avec usages et appels nommés du même module | Renommage public, intermodules, membres de classe, interfaces COM et callbacks absent ; compilation conditionnelle ou liaison ambiguë refusée |
| Navigation | Procédures, recherches, plages sélectionnées, symboles/déclarations, signets et navigation native | Index syntaxique : pas de résolution COM ni de variables implicites ; homonymes et SHA périmés refusés |
| Affichage / volets | Fenêtres, focus/visibilité, code/concepteur, défilement, sélection, fractionnement, fenêtres liées, arrangement et géométrie | Variantes DPI/écrans et premier ancrage à droite à qualifier ; fenêtres natives sans état accessible signalées explicitement |
| Affichage / explorateur de projets | Projets/composants/propriétés accessibles par VBIDE ; sélection/navigation via commandes | Vue exacte de l'arbre natif, recherche de nœuds et état de développement/repli des branches sans contrat spécifique |
| Affichage / explorateur d'objets | Bibliothèques/classes/membres, pagination, sélection exacte et description native | Variantes de langues et lecture SOLIDWORKS à qualifier ; observation UI ne prouve pas la résolution sémantique |
| Affichage / barres | Catalogue de commandes, placement, visibilité, commandes natives ajoutées/supprimées, profils SQLite par hôte | Persistance Excel qualifiée ; SOLIDWORKS et variantes de contextes à qualifier ; pas d'OnAction arbitraire |
| Insertion / code | Procédures Sub/Function/Property, modules/classes/forms, événements du concepteur, références | Usages avancés des interfaces et attributs persistés de classe à qualifier ; les setters ne simulent pas une interface COM |
| Format / géométrie | Positions/dimensions, alignement, tailles communes, centrage, distribution, espacement, grille, ordre de tabulation | Ajustement collectif au contenu et préférences complètes de grille absents ; certaines propriétés natives sont refusées après échecs observés |
| Format / ordre visuel | `z_order_form_control` avant/arrière et relecture de présence | Inventaire indépendant de l'ordre Z et déplacement d'un seul rang absents ; MSForms n'expose pas cet ordre via Controls, résultat explicitement non vérifié |
| Concepteur / contrôles | Types installés ; création racine/Frame/Page ; arbre et propriétés scalaires/objets ; polices/images ; renommage/suppression | Toute propriété exposée n'est pas forcément inscriptible ; propriétés dangereuses refusées ; effets/persistance/runtime de toutes les combinaisons à qualifier |
| Concepteur / conteneurs | Frame, MultiPage, Pages et TabStrip ; ajout/retrait pages/onglets et propriétés des nœuds | Contrôles tiers et combinaisons de conteneurs à qualifier ; groupement/dégroupement natif sans contrat spécifique |
| Concepteur / listes | Lignes paginées, mutations, initialisation multicolonne, liaison de listes | Combinaisons de contrôles/hôtes à qualifier ; garanties sur les liaisons liées au document limitées à l'état relu |
| Concepteur / copie | Sélection canonique ; copies typées ; presse-papiers natif ; récupération d'une coupe ; historique natif | Copie complète d'objets arbitraires et fidélité totale des images/événements/propriétés non garantie ; relecture des différences obligatoire |
| Concepteur / boîte à outils | Fenêtre et catalogue des types de contrôles ; ajout via contrat | Onglets/catégories, personnalisation des boutons et réorganisation de la boîte à outils absents ; le catalogue ActiveX n'en est pas un remplacement |
| Concepteur / exécution | Lancement UserForm, état asynchrone, observation des formes runtime et diagnostics | Cycle des événements et propriétés effectives à qualifier selon type/hôte ; livraison d'une commande différente de réussite runtime |
| Débogage / compilation | Compilation native, erreur sélectionnée, dialogue et options d'interruption | Variantes des diagnostics et langues à qualifier ; retour sans diagnostic observé distinct d'une garantie sur tous les scénarios |
| Débogage / commandes | Run/Break/Reset, pas à pas, pas principal/sortant, exécuter jusqu'au curseur, instruction suivante | Cas hôtes/dialogues particuliers à qualifier ; aucune génération d'inventaire fictif à partir des commandes disponibles |
| Débogage / points d'arrêt | Basculement exact, suppression globale, arrêt effectif qualifié | Inventaire indépendant des marqueurs et emplacement exact du pointeur d'exécution absents ; état d'un bouton de commande insuffisant |
| Débogage / variables | Fenêtres locales/espions, arbres, recherche/pagination, ajout/modification/retrait d'espion, Quick Watch | Dernier essai SOLIDWORKS sans lignes accessibles ; arbres COM et langues à qualifier ; ne pas présenter un arbre inaccessible comme vide |
| Exécution / procédures | Sub sans paramètre ; Sub/Function publique standard avec scalaires ; suivi asynchrone ; arguments nommés et omission d'optionnels | Objets/tableaux, retour COM structuré et appels de membres absents ; appels nommés ParamArray/conditionnels refusés ; erreurs/modalités peuvent dépasser la livraison |
| Exécution / volet Exécution | Instructions natives, texte relu et dialogues | Contenu UI borné ; pas de valeurs runtime inventées ni de relance automatique d'une commande encore en cours |
| Outils / références | Lire bibliothèques/types/membres ; ajout par GUID/fichier et retrait sous version | Variantes COM installées, références cassées et autres hôtes à qualifier ; aucun contournement de la politique native de confiance |
| Outils / projet | Métadonnées, propriétés scalaires, Instancing des classes, renommage du projet Excel enregistré | Configuration de protection par mot de passe absente ; projet protégé lisible seulement selon API ; renommage hors périmètre Excel refusé |
| Outils / signature | Certificats, présence d'une signature, dialogue natif et ajout ; confiance du certificat ; vérification du fichier Office via SIP VBA Microsoft | SIP dédié requis pour le digest : enregistrement temporaire soumis à autorisation, qualification native NOT_RUN ; `.swp`/Access hors formats pris en charge ; certificat/horodatage détaillés non extraits |
| Outils / options | Lecture native des onglets ; mutation bornée Éditeur/Général FR/EN, versions et relecture | Mutation Format de l'éditeur/Ancrage absente ; autres langues à qualifier ; thème natif VBE exclu du périmètre retenu |
| Outils / macros | Lancement par identité VBIDE inspectée | Catalogue et pilotage spécifique du sélecteur natif des macros absents |
| Compléments | Liste et changement de connexion d'un complément sous version | Connexion/déconnexion de compléments tiers et contextes hôtes à qualifier |
| Fenêtre / historiques | Disposition, fenêtres liées et historique natif de code | Historique natif de code limité à un projet ouvert sans UserForm ; garde maintenue car la pile partagée peut affecter un autre document |
| Événements / synchronisation | Hooks projets/composants/références, observation du texte et états de session | Rechargement de composants et changements externes complexes à qualifier ; un événement absent ne prouve pas l'absence de changement |
| Aide | Métadonnées HelpFile/HelpContextID et descriptions de l'explorateur | Ouverture/navigation contextuelle de l'aide native sans contrat spécifique |

## Priorité de développement

1. Finaliser les contrats modules/classes et exécution avec liaison vérifiable : renommage borné, arguments nommés, références et persistance.
2. Qualifier le digest de signature Office après autorisation du SIP ; conserver une réponse « vérificateur indisponible » autrement.
3. Développer la configuration du projet/protection et les surfaces natives de boîte à outils/options encore absentes.
4. Explorer les marqueurs du débogueur et les arbres SOLIDWORKS, sans réduire les gardes ni supposer une API inexistante.
5. Compléter fichier/projets, explorateur de projets, aide et qualification des variantes de concepteur.

Les scénarios natifs et leurs preuves sont dans [Qualification native](native-qualification.md). Les paramètres et permissions des outils sont dans [le catalogue LLM](vbe-tools.md). Les limites de chaque nouveau contrat sont dans [Extensions fonctionnelles](functional-extensions.md).
