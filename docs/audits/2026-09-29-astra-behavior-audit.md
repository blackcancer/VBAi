# Revue de comportement Astra — 29 septembre 2026

## Périmètre et limites

Revue en lecture seule par **GPT-6 Astra, raisonnement high**, du commit `8846c38`, après intégration et publication de la PR #12. Les défauts sont confirmés par leurs chemins de code ; la revue elle-même n’a ni manipulé les hôtes VBE ni qualifié leur manifestation dans Excel ou SOLIDWORKS. Les anciens assemblages Release datés du 27 septembre n’ont pas été utilisés comme preuve du code courant.

Le chantier de couverture est arrêté à réception du rapport, conformément à la demande. Les deux lots déjà terminés comptent **20 tests réussis**, dont **six scénarios ajoutés**, sans échec. Preuves : `artifacts/coverage-during-audit/first/coverage.trx` et `second/coverage.trx`. Aucun nouveau passage global de couverture n’est lancé après le rapport. Les validations des corrections ci-dessous sont des régressions fonctionnelles distinctes.

## Constats et corrections appliquées

| Priorité | Constat confirmé | Correction |
| --- | --- | --- |
| P2 | L’ouverture des sélecteurs du chat impose 36 pixels après mise à l’échelle, indépendamment des tailles de contrôle et de police. | La hauteur utilise les tailles actuelles et préférées des sélecteurs, leurs marges et les espacements du conteneur. La barre Monaco utilise aussi les dimensions naturelles des commandes, avec place pour le défilement horizontal. |
| P2 | Une session restaurée conserve ses activités `inProgress`, affichées actives alors qu’aucune intervention ne continue. | À l’activation d’une session inactive, ces activités deviennent `interrupted`, sans inventer de résultat ni de durée. Le libellé « Interrupted » est ajouté aux treize catalogues ; les étapes sont repliées par défaut. |
| P2 | La sélection d’un autre document ferme le diff JavaScript sans remettre `showingDiff` à faux dans l’hôte. | Une sélection commune synchronise Monaco et l’hôte lors des changements d’onglet, réouvertures, navigation par outils et initialisation. L’ouverture d’un nouveau document remet aussi l’hôte en mode code. Les révisions examinées sont conservées. |
| P2 | Les choix « Use edited version » et « Reload VBA version » deviennent des icônes seules, contrairement au contrat de l’interface compacte. | Les deux commandes conservent leur icône, leur texte visible et leur taille automatique. Le défilement horizontal permet les traductions longues. |

Sources : [chat](../../src/VBAi/Llm/Chat/ChatWindow.cs), [sessions](../../src/VBAi/Llm/Chat/ChatWindow.Sessions.cs), [activités](../../src/VBAi/Llm/Chat/ChatWindow.Activities.cs), [Monaco](../../src/VBAi/Editor/ModernEditorWindow.cs), [outils Monaco](../../src/VBAi/Editor/ModernEditorWindow.Tools.cs), [Designer Monaco](../../src/VBAi/Editor/ModernEditorWindow.Designer.cs).

## Optimisation des fragments d’activité

Le chemin précédent retire puis réinsère le propriétaire du groupe à chaque fragment, reconstruisant les contrôles de toutes ses étapes. Un fragment ajouté dont le type, le titre, le statut et la durée restent identiques met désormais à jour directement le champ de texte matérialisé. Les ajouts d’étapes, changements de métadonnées et vues absentes ou libérées conservent le parcours de reconstruction.

La régression envoie **300 fragments** dans un groupe de plusieurs étapes : aucun retrait/réinsertion du groupe, identité du champ conservée et état d’expansion de l’étape précédente préservé. Le changement de statut terminal déclenche encore le rafraîchissement nécessaire. Cela vérifie la réduction du travail structurel ; **aucun gain de temps chiffré n’est revendiqué**. Une cadence de rafraîchissement supplémentaire demanderait une mesure de performance avant introduction.

## Qualification

Le premier lot de régressions fonctionnelles compte **8 réussites, 0 échec** : restauration et libellé des activités, fragments, expansion des sélecteurs après mises à l’échelle 100/150/200 % et changement de police, choix de conflit et barre dimensionnée, parcours réel WebView2 « comparaison → onglet → réouverture → nouveau module ». Les facteurs sont exercés par la mise à l’échelle WinForms ; ils ne constituent pas une qualification sur trois moniteurs de DPI différents.

Preuve : `artifacts/astra-bug-fixes/compiled/regressions.trx`. Le lot élargi des surfaces affectées compte **178 réussites, 0 échec, 0 ignoré** dans `artifacts/astra-bug-fixes/accepted/regressions.trx`. Le dernier contrôle de fermeture compte **2 réussites, 0 échec** dans `artifacts/astra-bug-fixes/close/regressions.trx`, avec fermeture du dernier onglet comparé dans WebView2 et conservation de la cible pendant un changement de sélection. Les builds Debug et Release réussissent sans avertissement ni erreur ; **46 concepteurs WinForms** passent le chargement, le redimensionnement et l’accès aux composants enfants (`artifacts/astra-bug-fixes/designers/`).

Aucune macro utilisateur n’est exécutée ou enregistrée. SOLIDWORKS reste non qualifié pour ce lot. La documentation XML couvre **5 353 / 5 353 déclarations** ; l’organisation conserve **240 miroirs pour 297 sources**.
