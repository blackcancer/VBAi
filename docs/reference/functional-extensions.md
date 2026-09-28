# Extensions fonctionnelles VBE

Synthèse des extensions intégrées à `main` et du lot complémentaire du 28 septembre 2026. Les fonctions ci-dessous sont exposées au pont et au modèle ; aucune structure WinForms n’a été modifiée dans ces lots IDE. Le thème natif VBE demeure exclu.

## Fonctions ajoutées ou étendues

### Classes, arguments variables et préférences natives

- `preview_class_member_rename` / `apply_class_member_rename` : membre explicitement privé d'une classe ordinaire ; Sub/Function ou famille Property Get/Let/Set entière. SHA du module, version du projet complet et export `.cls` natif contrôlés avant mutation. Les attributs cachés des membres, classes exposées/prédéclarées, interfaces, callbacks et liaisons ambiguës sont refusés. Diff de conversation et annulation par module ; aucune transaction globale ni compilation implicite prétendue.
- `run_procedure_values` accepte désormais un dernier `ParamArray` Variant, explicite ou implicite : zéro à trente arguments positionnels au total, avec préfixes ByVal obligatoires. Un tableau transmis reste un argument distinct. Les arguments nommés et les préfixes Optional sont refusés pour cette signature, conformément aux contraintes VBA ; le transport reste limité à Excel.
- `read_vbe_options` expose les choix natifs de police et de palettes, ainsi que les trois palettes de chaque catégorie de couleurs dans `FormatCategories`. La version protège toutes les catégories. `set_vbe_option.Query` désigne une catégorie native exacte pour cette mutation. `NativeIndex:n` représente une entrée sans libellé, sans inventer de couleur RGB. Une taille lisible avec catalogue natif vide reste non modifiable.

Les preuves hôtes de ce lot sont suivies séparément dans [Qualification native](native-qualification.md). Une matrice locale verte ne prouve pas la réussite dans Excel ou SOLIDWORKS.

Contraintes VBA : Microsoft impose des [arguments positionnels avec ParamArray](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/a-procedure-with-a-paramarray-argument-cannot-be-called-with-named-arguments) et interdit la [combinaison Optional/ParamArray](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/can-t-have-paramarrays-with-optional-arguments).

### Lot de complétion IDE du 28 septembre

- `preview_procedure_rename` / `apply_procedure_rename` : renommage standard intermodules, contrôle de tout le projet avant écriture, diff et historique séparés par module. Les erreurs partielles sont relues ; aucune transaction globale n'est prétendue.
- `read_navigation_surface` / `change_navigation_surface` : explorateur de projets et sélecteur de macros accessibles ; fallback MSAA pour lire les pages Toolbox. Tokens natifs et version de la surface ; sélection Toolbox observée séparément des actions qualifiées, boutons non exposés explicitement signalés ; aucune action par coordonnées ou raccourcis.
- `run_procedure_values` / `procedure_values_status` : invocation unique via Excel.Application.Run, arguments Variant/tableaux rang 1/2, retour JSON avec bornes et longueurs natives. Exige un classeur sauvegardé associé par identité COM/PID/chemin. Aucun helper VBA ni deuxième appel pour récupérer une valeur. Les tableaux typés/ByRef, classes, objets COM et dates sont refusés ; ParamArray est décrit dans le lot complémentaire ci-dessus.
- `list_macros` : catalogue syntaxique public des modules standard. Les candidats natifs sont distingués des fonctions, des arguments et des exclusions Option Private Module/compilation conditionnelle.
- `project_collection_state`, `create_standalone_project`, `open_standalone_project`, `close_standalone_project` : API VBProjects Add/Open/Remove, gardes sur tous les projets, fermeture de SWP déjà sauvegardés seulement. Qualification SOLIDWORKS NOT_RUN.
- `preview_fit_form_content` / `apply_fit_form_content` : enfants directs, conteneur ou étendue de défilement. Les dimensions de conception racine utilisent VBComponent.Properties, celles du viewport le Designer. Arrondi supérieur au pixel et relecture des dimensions/enfants ; toute dérive après écriture est signalée incertaine.
- `read_project_protection` / `set_project_protection` : dialogue du projet exact, fichier secret UTF8 local pour verrouillage, aucune restitution du mot de passe. La persistance nécessite sauvegarde/réouverture et n'est pas affirmée par la fermeture du dialogue.
- `open_native_ide_dialog` : ouverture seule de Macros, Impression, Aide ou Contrôles supplémentaires. Aucune soumission d'impression ni installation implicite de contrôle.
- `open_project_help` : CHM local et HelpContextID natif ; le sujet affiché demeure non vérifié.
- Sauvegarde Word/PowerPoint : identité du VBProject et PID du document ; formats DOCM/DOTM/PPTM/POTM/PPSM, sans écrasement lors de la première sauvegarde. Qualification réelle de ces hôtes NOT_RUN.
- Options Éditeur/Format/Ancrage/Général : valeurs bornées et choix exacts accessibles ; les palettes non lisibles restent refusées.

Les points d'arrêt et le pointeur d'exécution ne sont pas inventoriables par les attributs TextPattern testés dans Excel : BackgroundColorAttribute reste NotSupported avant/après un point connu et en pause. Le groupement du Designer n'expose pas son appartenance dans la typelib MSForms inspectée ; les booléens CanUndo/CanRedo ne prouvent pas quelle action serait annulée. Ces manques restent explicites.

### Contrats précédents

Ce tableau conserve le périmètre initial du lot basé sur `b843972f`. Les extensions décrites ci-dessus ajoutent notamment les préférences Format/Ancrage et les adaptateurs de sauvegarde Word/PowerPoint ; leurs limites actuelles figurent dans l'inventaire fonctionnel.

| Surface | Contrat disponible | Limite exacte |
| --- | --- | --- |
| Procédures paramétrées | `run_procedure` puis `procedure_run_status` ; Sub/Function publique de module standard, SHA et mode conception, jusqu’à 30 arguments JSON scalaires ; `ArgumentNames` facultatif associe chaque valeur à un paramètre vivant et permet d'omettre les optionnels | Transmission native asynchrone au volet Exécution ; aucun objet, expression arbitraire ni valeur de retour structurée. Appels nommés ParamArray/conditionnels refusés. Une livraison ne prouve pas la réussite runtime. Refus des noms de projets homonymes pour l’appel qualifié. |
| Déclarations | `project_symbols` et références chat : variables, constantes, paramètres, types/champs, enums/membres, événements et Declare, avec portée et coordonnées | Index syntaxique ; aucune résolution COM, variable implicite ou analyse de liaison complète. Les références de déclaration refusent un SHA périmé. |
| Renommage local | `preview_local_rename` et `apply_local_rename` ; variable/constante explicitement déclarée dans une procédure, avec SHA, coordonnées et contrôle des collisions ; diff et journal undo/redo | Symboles de projet non pris en charge. Code conditionnel, DefType et ambiguïtés refusés. Membres qualifiés, types, labels, arguments nommés, commentaires et chaînes conservés. |
| Renommage de paramètre | `preview_parameter_rename` et `apply_parameter_rename` ; Sub/Function privée de module standard, SHA, procédure et coordonnées exactes ; déclaration, usages locaux et appels nommés directs/qualifiés par le module mis à jour ; diff LLM et annulation | Classes, formulaires, interfaces publiques/callbacks, appels externes et compilation conditionnelle refusés. Ce contrat ne prétend pas effectuer une résolution sémantique complète du projet. |
| Digest VBA enregistré | `verify_vba_signature_file` ; fichier Office immuable pendant lecture, SHA-256 et WinVerifyTrust ciblant le SIP VBA ; résultat distingue valide/confiance, absence, digest altéré et vérificateur indisponible | Dépend du SIP VBA Microsoft dédié installé ; pas de substitution par la signature du document Office. Document fermé ou copie enregistrée séparée : un fichier ouvert pour écriture est refusé pour préserver une lecture immuable. Formats `.swp`/Access refusés. SIP natif NOT_RUN en attente d'autorisation ; erreurs de chaîne/confiance ne deviennent pas un diagnostic de digest valide. |
| Nom de projet | `set_project_property`, `Property=Name`, setter VBIDE typé et relecture du nom, chemin et SHA des modules | Excel uniquement : classeur enregistré, inscriptible, projet non protégé, ciblé par chemin absolu. Renommage de métadonnées ; références qualifiées et appelants externes inchangés. Pas de sauvegarde automatique. |
| Sauvegarde SWP | `project_persistence_status`, `save_host_document`, `save_host_document_as` pour VBProject standalone Type=101 | `.swp` uniquement ; chemin explicite, version de projet et mode conception. SaveAs n’écrase pas de fichier. Lecture native du chemin, Saved et fichier non vide ; pas de preuve de rechargement SOLIDWORKS. Les autres projets hôtes non Excel restent non pris en charge. |
| Barres | `toolbar_controls`, `create_toolbar`, `remove_toolbar`, `add_toolbar_command`, `remove_toolbar_command`, versions de collection/contrôles | Barres normales, commandes natives existantes avec ID/libellé exact ; aucun OnAction arbitraire. Suppression limitée aux boutons possédés par VBAi et aux barres VBAi vides. `Temporary` vaut true par défaut ; false enregistre un profil SQLite privé par hôte, restauré au chargement. Les commandes sont copiées depuis leur contrôle natif exact ; leur libellé peut varier selon le contexte. Persistance Excel qualifiée, SOLIDWORKS non qualifiée. |
| Options | `set_vbe_option` via InvokeAsync, avec version, onglet et libellé natifs exacts | Liste fermée Éditeur/Général en français/anglais : syntaxe, déclarations, complétion, aide/info-bulles, indentation, tabulations 1–32, compilation et interruption sur erreurs. Ni sécurité, thème, formatage ni ancrage. Relecture de la valeur avant OK ; réouverture requise pour prouver la persistance. |
| Confiance de certificat | `certificate_trust` avec empreinte exacte du magasin CurrentUser/My ; chaîne code-signing et politique Authenticode Windows | Cache uniquement, sans téléchargement ni clé privée. Révocation inconnue = indéterminé. Aucun digest VBA, liaison du signataire à la macro ou horodatage de signature vérifié. |

Le diagnostic cryptographique utilise les drapeaux de chaîne et de révocation en cache de [CertGetCertificateChain](https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-certgetcertificatechain). La sauvegarde SWP suit [VBProject.SaveAs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/saveas-method-vba-add-in-object-model), réservé aux projets standalone.

## Preuves

- Compilation de la solution : aucune erreur ni avertissement.
- Suite globale Unit : **905 réussis, 0 échec, 0 ignoré** ; rapport `artifacts/functional-extensions/tests/unit.trx`. Le pourcentage de couverture du bilan précédent n’a pas été remesuré pour ces ajouts.
- Scénario reproductible : `tools/probes/Test-FunctionalExtensionsExcel.ps1`, instance Excel jetable exclusivement, état AccessVBOM enregistré puis restauré.
- Excel : index de déclaration locale, renommage local appliqué, macro exécutée (résultat 7), annulation et code intégral relu identique ; création/ajout/retrait/suppression d’une barre temporaire.
- Excel : setter de nom de projet natif essayé sur six cycles ; service de production essayé avec deux projets homonymes ouverts. Nom restauré, chemin et code conservés. L’assertion de version utilise désormais l’identité du projet demandée, au lieu de son nom ambigu.
- Rapports locaux : `artifacts/functional-extensions/native/functional-extensions-excel.json`, `native-project-rename-trials.json`, `native-project-rename-service.json` et `security-before.json`.
- Les tests de sauvegarde standalone utilisent des doubles COM. SOLIDWORKS n’a pas été lancé.
- Qualifications depuis le complément Excel installé : arguments scalaires et résultat runtime relus, 13 essais de préférences réouverts puis restaurés, barre et commande restaurées après redémarrage. Voir le [bilan des cinq qualifications](native-qualification.md), qui détaille les cas validés et les cas restant non exécutés.

## Fonctions encore absentes ou bornées

1. Renommage des membres publics/Friend de classe, interfaces et références COM/externes : le nouveau contrat des membres explicitement privés reste conservateur, comme les renommages local, paramètre privé et procédure standard intermodules.
2. Inventaire indépendant exhaustif des breakpoints et du pointeur d’exécution : les commandes natives et l’état de pause ne constituent pas cet inventaire.
3. Historique natif partagé de code avec plusieurs projets ouverts ou des UserForms : gardes existantes conservées. Le journal VBAi et l’historique du Designer restent disponibles séparément.
4. Personnalisation des onglets et catégories de la boîte à outils native : non implémentée. Catalogue ActiveX et ajout de contrôles existants ne la remplacent pas.
5. Sauvegarde des autres hôtes : adaptateurs Word/PowerPoint implémentés mais non qualifiés nativement ; Access et les autres hôtes restent sans adaptateur.
6. Lecture des valeurs du débogueur SOLIDWORKS lorsque l’hôte n’expose aucune ligne accessible : pas de solution native qualifiée ; dernier essai zéro ligne.
7. Validation cryptographique de la signature VBA elle-même (digest, certificat signataire et timestamp) : le diagnostic de certificat ne la remplace pas.
8. Renommage des projets protégés, non enregistrés et autres hôtes : refus explicite conservé ; aucun contournement de protection.

Ces éléments sont distincts des qualifications encore requises pour les fonctions déjà implémentées : contrôles ActiveX tiers autorisés par l’hôte, propriétés/événements UserForms au-delà des cas mesurés, autres DPI et ancrage natif, autres arbres et langues du débogueur, SOLIDWORKS et intégrations authentifiées. Un tableau Excel de 1 000 éléments est désormais qualifié ; cela ne prouve pas les arbres de tous les objets COM.
