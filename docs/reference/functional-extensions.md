# Extensions fonctionnelles VBE

Développement du 28 septembre 2026 sur `codex/chat-ux`, basé sur `b843972f`. Les fonctions ci-dessous sont exposées au pont et au modèle ; aucune structure WinForms n’a été modifiée. Le thème natif VBE demeure exclu.

## Fonctions ajoutées ou étendues

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

1. Renommage sémantique entre modules, paramètres, membres de classe et références COM : le nouveau renommage est local et conservateur.
2. Inventaire indépendant exhaustif des breakpoints et du pointeur d’exécution : les commandes natives et l’état de pause ne constituent pas cet inventaire.
3. Historique natif partagé de code avec plusieurs projets ouverts ou des UserForms : gardes existantes conservées. Le journal VBAi et l’historique du Designer restent disponibles séparément.
4. Personnalisation des onglets et catégories de la boîte à outils native : non implémentée. Catalogue ActiveX et ajout de contrôles existants ne la remplacent pas.
5. Sauvegarde des projets hôtes autres qu’Excel et standalone `.swp`, dont documents Word/Access : non implémentée.
6. Lecture des valeurs du débogueur SOLIDWORKS lorsque l’hôte n’expose aucune ligne accessible : pas de solution native qualifiée ; dernier essai zéro ligne.
7. Validation cryptographique de la signature VBA elle-même (digest, certificat signataire et timestamp) : le diagnostic de certificat ne la remplace pas.
8. Renommage des projets protégés, non enregistrés et autres hôtes : refus explicite conservé ; aucun contournement de protection.

Ces éléments sont distincts des qualifications encore requises pour les fonctions déjà implémentées : contrôles ActiveX tiers autorisés par l’hôte, propriétés/événements UserForms au-delà des cas mesurés, autres DPI et ancrage natif, autres arbres et langues du débogueur, SOLIDWORKS et intégrations authentifiées. Un tableau Excel de 1 000 éléments est désormais qualifié ; cela ne prouve pas les arbres de tous les objets COM.
