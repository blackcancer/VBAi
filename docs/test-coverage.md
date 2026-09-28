# Couverture automatisée du complément

## Mesure globale à 100 % après extension IntelliSense Monaco

Mesure du **29 septembre 2026**, source `611f882`, après les références dynamiques, les métadonnées COM enrichies, les alias VBA et l’édition automatique. Un passage global final VSTest instrumenté, sans exclusion de production, fournit :

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 834 réussis, 0 échec, 20 ignorés**, 9 min 27 s |
| Lignes du complément C# | **29 712 / 29 712 — 100 %** |
| Branches du complément C# | **30 845 / 30 845 — 100 %** |
| Qualification ciblée finale | **10 réussis, 0 échec, 0 ignoré**, renderer WebView2, matrice JavaScript et parcours Excel réel |
| Matrice JavaScript | **57 scénarios réussis**, références, survols, signatures, blocs et formatage |
| Build de la solution | **0 erreur, 0 avertissement** |
| Concepteurs WinForms | **46 surfaces validées** |
| Métadonnées Visual Studio | **27 contrôles réussis** |
| Organisation miroir | **237 miroirs pour 294 sources**, scénarios complémentaires séparés |
| Documentation IntelliSense | **5 216 / 5 216 déclarations**, aucune erreur syntaxique |

Preuves : `artifacts/monaco-language/global-final-results/global.trx`, `36b2284b-9307-4796-9f94-028820a3b402/coverage.cobertura.xml` et `coverage.json` dans le même répertoire. Toutes les classes instrumentées atteignent 100 % en lignes et branches. La qualification ciblée est sous `artifacts/monaco-language/native-final-results/native.trx` ; les contrôles de structure et documentation sont sous `artifacts/monaco-language/`.

Deux passages antérieurs comportaient chacun un échec et ne servent pas de preuve de suite verte : chemin de cache Git trop long sous `global-build`, puis lecture du marqueur de connexion avant fermeture du fichier dans la fixture Copilot. La sortie courte et la publication atomique du marqueur corrigent ces causes. La mesure finale utilise ces corrections.

Les 20 scénarios conditionnels d’hôtes et de comptes restent désactivés dans cette mesure. Le parcours Excel séparé vérifie les références Excel/VBA, l’ajout/retrait Office et Scripting, les chaînes et collections, les survols, les blocs, le formatage et l’annulation, avec le module natif inchangé ; voir [le bilan natif](reference/native-qualification.md). Les compteurs globaux concernent l’assembly C# `CodexVBE` ; les 57 scénarios JavaScript constituent une qualification fonctionnelle distincte, sans prétendre mesurer ses lignes ou branches. Le moteur C++ et les combinaisons d’hôtes non exécutées conservent leur qualification propre.

## Mesure globale à 100 % après PR #11

Mesure du **29 septembre 2026**, source `b20b46f`, après fusion de la PR #11 et qualification de la file de messages ainsi que des vues WinForms du transcript. Le commit `0642e1c` suivant ne change que l’indentation des commentaires XML ; Roslyn confirme l’équivalence du code. Un seul passage global final VSTest instrumenté, sans exclusion de production, fournit les compteurs suivants :

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 828 réussis, 0 échec, 19 ignorés**, 9 min |
| Lignes du complément C# | **29 663 / 29 663 — 100 %** |
| Branches du complément C# | **30 793 / 30 793 — 100 %** |
| Build de la solution | **0 erreur, 0 avertissement** |
| Concepteurs WinForms | **46 surfaces et 409 contrôles enfants validés** |
| Métadonnées Visual Studio | **27 contrôles réussis** |
| Organisation miroir | **237 miroirs pour 294 sources**, scénarios complémentaires séparés |
| Documentation IntelliSense | **5 208 / 5 208 déclarations**, aucune erreur syntaxique ni différence de code Roslyn |
| Catalogues de langue | **14 catalogues, 763 textes obligatoires chacun**, aucune clé absente, vide ou dupliquée |

Preuves : `artifacts/pr11/global-final-results/global.trx`, `badef2be-8c80-465f-85be-c7ccdd1201a3/coverage.cobertura.xml` et `coverage.json` dans le même répertoire. Toutes les classes instrumentées atteignent 100 % en lignes et branches. Les autres preuves sont sous `artifacts/pr11/designers-qualified/`, `metadata-qualified/`, `layout.json` et `intellisense-audit.txt`.

La qualification a corrigé deux menus Designer sans nom et les onze catalogues auxquels manquaient les nouveaux libellés de la file. La surcharge `ChatTextContentView.Append` conserve ses appels existants tout en évitant le paramètre enum optionnel qui empêchait l’instrumentation Coverlet. Aucun filtre d’exclusion ni package de substitution n’a été ajouté. Le passage antérieur `global-qualified-results` comportait deux échecs et ne sert pas de preuve de suite verte.

Les 19 scénarios conditionnels d’hôtes et de comptes restent désactivés dans cette mesure. Aucun hôte utilisateur n’a été fermé ni sa DLL chargée remplacée pour cette intégration ; les constructions utilisent des sorties isolées. Les qualifications natives antérieures restent datées dans [le bilan natif](reference/native-qualification.md). La mesure concerne l’assembly C# `CodexVBE` ; le moteur C++ et toutes les combinaisons d’hôtes, de DPI et de contrôles tiers nécessitent leur propre qualification.

## Mesure globale à 100 % après PR #10

Mesure du **28 septembre 2026**, source `48dfa88`, après intégration de la PR #10 et des matrices complémentaires. Un seul passage global VSTest instrumenté, sans exclusion de production, fournit les deux compteurs exacts :

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 770 réussis, 0 échec, 19 ignorés**, 7 min 55 s |
| Lignes du complément C# | **28 987 / 28 987 — 100 %** |
| Branches du complément C# | **30 527 / 30 527 — 100 %** |
| Build de la solution | **0 erreur, 0 avertissement** |
| Concepteurs WinForms | **33 surfaces validées** |
| Organisation miroir | **208 miroirs pour 265 sources**, scénarios complémentaires séparés |

Preuves : `artifacts/cov/global-qualified-100-results/global.trx`, `ea683f23-fd08-461a-8e41-513991d2fe8e/coverage.cobertura.xml` et `coverage.json` dans le même répertoire. Les compteurs couverts et totaux sont égaux ; aucune classe instrumentée ne reste sous 100 %. Le passage `global-100` interrompu est obsolète et ne sert pas de preuve. Les anciens pourcentages ci-dessous sont historiques.

Les 19 scénarios conditionnels nécessitant un hôte ou un compte connecté ne sont pas activés dans cette mesure. Les essais réels exécutés séparément comprennent cinq scénarios SOLIDWORKS (module, classe, formulaire, exécution et breakpoint), son test VSTest de connexion, deux tests fournisseurs connectés GitHub/Codex et le démarrage Monaco dans Excel. Leur détail et leurs limites restent dans [la qualification native](reference/native-qualification.md). Ces passages séparés ne sont pas ajoutés aux compteurs globaux.

La couverture mesurée concerne l'assembly C# `CodexVBE`. Elle vérifie l'exécution des lignes et branches de ses contrats automatisés ; elle ne constitue pas une qualification universelle des combinaisons Office, COM, DPI, signatures et contrôles tiers. Le moteur C++ n'est pas instrumenté par Coverlet.

## Après intégration de la PR #9 — diagnostics et attributs Monaco

Mesure locale du **28 septembre 2026**, fusion de `d91ffb8` et garde supplémentaire du renommage concurrent. Construction isolée : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 388 réussis, 0 échec, 15 ignorés**, 6 min 25 s |
| Lignes du complément C# | **25 883 / 27 930 — 92,67 %** |
| Branches du complément C# | **26 501 / 29 523 — 89,76 %** |
| Excel natif Monaco, passage séparé | **1 réussi, 0 échec, 0 ignoré** |
| Concepteurs WinForms | **32 surfaces validées** |
| Organisation miroir | **190 miroirs pour 254 fichiers de production** |
| Catalogue LLM | **204 outils** |

Il reste **2 047 lignes et 3 022 branches** C# non exécutées. Aucun code de production du complément n'est exclu. Les nouvelles opérations de remplacement ajoutent des chemins à couvrir ; l'objectif 100 % reste non atteint. Les 14 scénarios Excel et le scénario SOLIDWORKS sont désactivés dans la mesure globale. Le test Excel séparé ne contribue pas à ces pourcentages.

Preuves finales : `artifacts/pr9-integration/qualified-global-final/global.trx`, `116dfc83-d9d5-4558-8e15-e3e6dfaf0e09/coverage.cobertura.xml`, `coverage-summary.json` et `coverage-inventory.csv` dans le même répertoire. Le rapport vide du premier passage exploratoire n'est pas une mesure de couverture. Le lot ciblé précédent compte **113 réussis, 0 échec, 1 ignoré** (`artifacts/pr9-integration/contracts/contracts.trx`) ; le correctif ultérieur est validé dans la suite finale et dans `artifacts/pr9-integration/native-final/excel.trx`. Les limites natives sont précisées dans [la qualification](reference/native-qualification.md).

## Après intégration de la PR #8 — thème natif expérimental

Mesure locale du **28 septembre 2026**, après Monaco et le thème natif jusqu'à `264e432`, avec les correctifs de fusion et d'arrêt. Build isolé : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 383 réussis, 0 échec, 15 ignorés**, 6 min 23 s |
| Lignes du complément C# | **25 782 / 27 714 — 93,03 %** |
| Branches du complément C# | **26 398 / 29 211 — 90,37 %** |
| Concepteurs WinForms | **32 surfaces validées** |
| Organisation miroir | **189 miroirs pour 253 fichiers de production** |
| Catalogue LLM | **204 outils** |
| Contrats ciblés thème/paramètres | **49 réussis, 0 échec, 0 ignoré** |
| Moteur C++ | **20 cycles synthétiques réussis**, chargement/hash/ABI validés sans hooks Office |

Il reste **1 932 lignes et 2 813 branches** C# non exécutées. Aucun code de production C# du complément n'est exclu. Le moteur C++ n'est pas instrumenté par Coverlet : ses tests synthétiques ne constituent pas une mesure de lignes/branches natives. Les pourcentages antérieurs sont historiques et l'objectif 100 % reste non atteint.

Les 14 scénarios Excel restent désactivés pendant l'utilisation concurrente d'Excel ; SOLIDWORKS reste NOT_RUN. Les preuves visuelles natives de la branche auteur sont conservées comme telles, sans annoncer leur répétition sur main. Les limites du thème et les correctifs d'intégration sont dans [le bilan natif](native-dark-theme.md).

Preuves : `artifacts/pr8-integration/qualified-global-final/global.trx`, `8c12c908-36f4-43cc-a5da-723cc8e9a251/coverage.cobertura.xml`, `coverage-summary.json` et `coverage-inventory.csv` dans le même répertoire ; tests ciblés `artifacts/pr8-integration/contracts/theme.trx` ; designers `artifacts/pr8-integration/designers/designers.json`.

## Après intégration de la PR #7 Monaco

Mesure locale du **28 septembre 2026**, branche Monaco intégrée jusqu'à `a4e3560`, avec les contrats IDE existants et le correctif de fermeture WinForms. Build isolé : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 358 réussis, 0 échec, 15 ignorés**, 6 min 1 s |
| Lignes du complément | **25 249 / 26 359 — 95,78 %** |
| Branches du complément | **25 995 / 27 897 — 93,18 %** |
| Concepteurs WinForms | **32 surfaces validées** |
| Organisation miroir | **184 miroirs pour 246 fichiers de production** |
| Catalogue LLM | **204 outils** |

Il reste **1 110 lignes et 1 902 branches** non exécutées. Aucun code de production n'est exclu. Les nouveaux adaptateurs Monaco ajoutent des chemins à couvrir ; les pourcentages ci-dessous sont historiques. Les 14 scénarios Excel (dont Monaco) sont désactivés pour éviter les essais concurrents ; le scénario SOLIDWORKS reste NOT_RUN. Les tests du véritable WebView2 passent dans cette suite, y compris les fermetures pendant l'initialisation et la création d'un contrôle d'état, avec conservation des brouillons sans écriture VBA.

Preuves : `artifacts/pr7-integration/qualified-global/global.trx`, `8589ae6a-88d1-4b50-b986-c338cffd0f3a/coverage.cobertura.xml`, `coverage-summary.json` et `coverage-inventory.csv` dans le même répertoire. Les **93 tests ciblés réussis** précèdent le dernier delta de la branche, qui est couvert par ce passage global. Les mesures de la branche Monaco et les limites natives sont distinguées dans [la documentation de l'éditeur](modern-editor.md).

## Lot complémentaire en qualification

Les contrats ParamArray, renommage de membres privés de classe et options natives ajoutent du code après la mesure ci-dessous. Le build du projet de tests passe avec **0 erreur et 0 avertissement**. Après correction des fixtures pour utiliser un véritable dialogue Win32 #32770, leur matrice locale donne **127 réussis, 0 échec, 0 ignoré** (`artifacts/vbe-next/contracts-qualified-final/contracts.trx`). Les gardes de production restent inchangées. Ce lot ciblé ne constitue pas un bilan global. La structure conserve **174 miroirs pour 231 fichiers de production**.

Le passage global local final donne **1 318 réussis, 0 échec, 14 ignorés**, en 5 min 48 s (`artifacts/vbe-next/qualified-local-final/global.trx`). Couverture du complément : **24 303 / 24 960 lignes — 97,36 %**, **25 316 / 26 377 branches — 95,97 %**. Aucun code de production n'est exclu. Les 13 scénarios Excel sont désactivés pour éviter les fermetures concurrentes ; SOLIDWORKS reste NOT_RUN. Il manque **657 lignes et 1 061 branches**. Les trois nouveaux fichiers de classes et ParamArray sont à **100 % lignes et branches** ; quatre méthodes du fichier des options restent incomplètes.

Preuves : `2e065eb0-4302-43c8-b049-14f8fb397b11/coverage.cobertura.xml`, `coverage-summary.json` et `coverage-inventory.csv` dans le même répertoire final. Les passages exploratoires échoués ne remplacent pas ce résultat. Les nouveaux scénarios Excel et leurs interruptions sont décrits dans la [qualification native](reference/native-qualification.md). La couverture actuelle n'est pas annoncée à 100 % et la mesure après PR #6 reste un résultat historique.

## Mesure après intégration de la PR #6 et complétion IDE

Mesure du **28 septembre 2026**, après la fusion `c5f64eb` et le lot IDE décrit dans l'[inventaire](reference/vbe-capability-inventory.md). Compilation : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 281 réussis, 0 échec, 1 ignoré**, 8 min 7 s |
| Lignes du complément | **23 821 / 24 469 — 97,35 %** |
| Branches du complément | **24 724 / 25 769 — 95,94 %** |
| Excel automatisé | **Dix scénarios réussis** |
| SOLIDWORKS | **NOT_RUN**, aucune instance préchargée |
| Concepteurs WinForms | **31 surfaces validées** après intégration PR #6 |
| Organisation miroir | **171 miroirs pour 228 fichiers de production** |
| Catalogue LLM | **197 outils**, dont 17 nouveaux contrats IDE |

Il reste **648 lignes et 1 045 branches** non exécutées. Les nouveaux adaptateurs natifs et les mises à jour GitHub ajoutent du code à couvrir ; les mesures historiques à 100 % ne décrivent pas le code actuel. Aucun code du complément n'a été exclu. La couverture mesure l'assembly `CodexVBE` dans VSTest ; les qualifications Excel sont indépendantes, et l'installation réelle par `VBAi.Updater` reste NOT_RUN faute de release signée de test.

Les dix scénarios Excel incluent la protection sauvegardée/réouverte, le renommage public intermodules et son annulation, l'ajustement UserForm, l'explorateur, les options Format/Ancrage restaurées, les pages Toolbox MSAA et les valeurs/tableaux retournés par une invocation unique. La [qualification native](reference/native-qualification.md) précise leurs limites.

Preuves : `artifacts/vbe-completion/qualified-global/global.trx` et `136ec1d1-d534-432d-918d-64b5d9210601/coverage.cobertura.xml` dans ce même répertoire. `coverage-summary.json` et `coverage-inventory.csv` contiennent les compteurs et méthodes encore incomplètes. Organisation : `artifacts/vbe-completion/test-layout.json`. Designers : `artifacts/pr6-integration/designers/designers.json`. Les passages exploratoires échoués sont conservés séparément et ne constituent pas la validation finale.

## Mesure après intégration de la PR #5

Mesure du **28 septembre 2026**, code **`a151498`** : À propos et rapports de problème GitHub/Outlook intégrés avec les fonctionnalités IDE de main. Compilation : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Suite globale VSTest | **1 189 réussis, 0 échec, 1 ignoré**, 6 min 34 s |
| Lignes | **21 578 / 21 658 — 99,63 %** |
| Branches | **22 462 / 22 545 — 99,63 %** |
| Excel automatisé | **Trois scénarios réussis** |
| SOLIDWORKS | **NOT_RUN**, aucune instance préchargée |
| Concepteurs WinForms | **29 surfaces validées** |
| Organisation miroir | **146 miroirs pour 200 fichiers de production** |

La PR ajoute du code dont la couverture reste à compléter : **80 lignes et 83 branches** non exécutées. Les 100 % du passage précédent ne sont pas la mesure actuelle. Aucun code de production n'a été exclu. Les publications GitHub et les livraisons Outlook sont simulées ; aucun rapport réel n'a été envoyé et ces parcours natifs ne sont pas qualifiés par cette suite.

Preuves locales : `artifacts/pr5-integration/results/global.trx` et `artifacts/pr5-integration/results/9dc1e6ae-ad7b-4920-aac7-6046ca8aefd5/coverage.cobertura.xml`. Les contrôles Designer et miroir sont sous `artifacts/pr5-integration/`.

## Mesure globale après intégration de la PR #4 et du lot IDE

Mesure du **28 septembre 2026**, code **`27389a8`**, comprenant la PR WinForms `codex/chat-ux`, le renommage de paramètres, les appels nommés et le vérificateur de signatures. Compilation de la solution : **0 erreur, 0 avertissement**.

| Mesure | Résultat |
| --- | --- |
| Lignes exécutables | **20 639 / 20 639 — 100 %** |
| Branches | **22 223 / 22 223 — 100 %** |
| Classes instrumentées incomplètes | **0** |
| Suite globale VSTest | **1 172 réussis, 0 échec, 1 ignoré** ; 6 min 37 s |
| Excel automatisé | **Trois scénarios réussis** dans la suite globale |
| SOLIDWORKS | **NOT_RUN** : aucune instance utilisateur préchargée |
| Organisation | **141 miroirs pour 193 sources**, scénarios et fixtures complémentaires |
| Concepteurs WinForms | **27 DesignSurface** chargées, redimensionnées et éditées ; contrôles enfants sélectionnables |
| Catalogue LLM | **180 outils** extraits de l'assembly reconstruite |

## Scénarios du lot

- Renommage de paramètre privé standard : déclaration, usages, appels nommés directs/qualifiés et imbriqués, continuations, collisions, compilation conditionnelle, SHA/mode, diff LLM et undo/redo.
- Exécution publique avec noms de paramètres vérifiés, réordonnancement et omission d'optionnels ; requêtes/noms/valeurs capturés avant la livraison ; gardes de taille, type, identité et syntaxe.
- Vérificateur Office : formats, disponibilité du SIP VBA, SHA et verrou de lecture, résultats HRESULT distincts, relecture du registre sans écriture et nettoyage WinVerifyTrust. Les résultats cryptographiques sont simulés localement ; leur qualification native demeure NOT_RUN tant que le SIP dédié n'est pas autorisé.
- PR WinForms : propriétés et marges du moteur de saisie, sérialisation/reset, apparence, événements avec/sans abonnés, ressources Designer, et libération des diff natifs lors du recyclage du transcript.
- Contrat du catalogue LLM : tous les champs requis/facultatifs, types et modes ; `ArgumentNames` validé comme tableau, sans modifier les politiques d'approbation.

Excel vérifie indépendamment les cellules **42** et **named call** après renommage et appel nommé, puis la restauration exacte du code. Le document enregistré est copié pour la vérification : Excel garde un verrou d'écriture sur son fichier ouvert. Le SIP absent produit **VerifierUnavailable**, pas une signature valide. Les deux autres scénarios vérifient le chargement/pont et la sauvegarde en présence d'une autre instance Excel jetable, laissée inchangée.

## Preuves et limites

- Rapport global : `artifacts/pr4-integration/final-qualified/a7899ebd-1a42-41ce-9820-9ebf2faa2aa3/coverage.cobertura.xml` ; résumé `coverage-summary.json` adjacent.
- Résultats globaux : `artifacts/pr4-integration/final-qualified/global.trx`.
- Organisation miroir : `artifacts/pr4-integration/test-layout.json`, produite par `tools/tests/Test-TestLayout.ps1`.
- Scripts Designer : `tools/tests/Test-WinFormsDesigners.ps1` et `tools/tests/Test-ChatDesigner.ps1`, exécutés sur l'assembly de la fusion.
- [Mesure précédente à 100 %](archive/test-coverage-pre-winforms-pr4.md) : historique, distinct de ce passage.

Aucune source du complément n'est exclue ; `ProviderTests.exe` est le seul exécutable de simulation exclu. Le collecteur mesure le processus VSTest, pas le code exécuté dans Excel ou SOLIDWORKS. Les doubles natifs vérifient les branches et erreurs du code ; 100 % de couverture ne qualifie pas toutes les fonctions dans tous les hôtes.

L'[inventaire fonctionnel complet du VBE](reference/vbe-capability-inventory.md) précise les fonctions absentes ou partielles. Les [qualifications natives](reference/native-qualification.md) distinguent les preuves Excel, les essais historiques SOLIDWORKS et le digest VBA non qualifié. L'audit IntelliSense avec Luna reste indépendant et non intégré à cette mesure.

## Reproduire la mesure

Voir [le projet de tests](../tests/README.md) pour VSTest, la couverture et les hôtes opt-in. Les rapports détaillés restent locaux sous `artifacts/`, ignoré par Git.
