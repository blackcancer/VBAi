# Matrice préparée avant exécution

| Surface | Cas |
| --- | --- |
| Déclaration | Private Sub et Function standard, ByVal/ByRef, Optional, tableau/ParamArray, suffixe, casse et continuation ; coordonnées et procédure exactes ; public, Property, classe/formulaire refusés |
| Usages | Déclaration et corps ; autres procédures, types, membres, labels, commentaires et chaînes conservés ; nouveau nom réservé/collision refusé |
| Appels | Directs, Call, sans parenthèses, Function dans expression, qualification du module, récursion, appels imbriqués, If Then/Else, arguments nommés et positionnels ; homonymes de membres préservés |
| Liaison | Noms de module/procédure masqués, signature ambiguë, directives conditionnelles et qualification de projet refusés |
| Service | Lecture/identité/type/catalogue/SHA absents ou périmés, ProcKind exact, mode conception, aperçu sans écriture, édition avec relecture, échec d'écriture/relecture, undo/redo |
| Modèle | Schéma requis, aperçu en Discussion, mutation en Agent/Automatique, blocage Lecture seule et projet étranger, diff et historique |

Les essais Excel utilisent un classeur jetable visible ; compilation et résultat runtime sont relus après renommage, puis le texte est restauré par undo.
