# Qualification de VBAi dans les hôtes Office

## Périmètre

Le lot du **29 septembre 2026** complète la suite globale passée après le renommage (**1 876 réussites, 0 échec, 21 ignorés**). Il exerce la DLL installée de `main`, MVID `acdda69f-d581-433a-9351-b9312736e9ce`, dans les applications Office 16 x64 disponibles sur ce poste. Chaque hôte possède une fenêtre visible et un fichier jetable. Le test vérifie le PID et le MVID réellement servis par le pont VBAi.

Les nouveaux tests sont dans `tests/VBAi.Tests/Integration/Hosts/Office/` ; le cycle de vie partagé est dans `Infrastructure/Hosts/OfficeVbeFixture.cs`. Ils sont conditionnels et ne remplacent pas les tests miroir de production.

## Résultats et limites

Le lot final compte **4 réussites, 0 échec, 1 ignoré**, en **53 secondes**. Chaque réussite documentaire regroupe les sept scénarios ci-dessous. Le test Outlook ignoré correspond à un prérequis absent, pas à une validation de son fonctionnement.

| Hôte | Qualification | Sauvegarde par la commande VBAi |
| --- | --- | --- |
| Word | Projet `.docm` jetable ; aucun changement à `Normal` | Refus vérifié : l'accès programmatique au projet n'est pas approuvé sur ce poste |
| PowerPoint | Présentation `.pptm` jetable | Défaut observé : l'adaptateur essaie de lire `Application.HWND`, membre indisponible à l'exécution |
| Access | Base `.accdb` jetable, y compris UserForm MSForms | Adaptateur absent ; refus explicite vérifié |
| Publisher | Publication `.pub` jetable | Adaptateur absent ; refus explicite vérifié |
| Outlook classique | **BLOCKED** : assistant de première configuration, aucun profil Office 16 configuré | Projet utilisateur non modifié ; parcours préparé en lecture seule |
| Visio | **NOT_INSTALLED** sur ce poste | Non testé |

Pour les quatre hôtes documentaires, le lot comprend sept scénarios :

1. Projet, environnement VBE, références et état de débogage.
2. Inspection de la bibliothèque propre à l'hôte ; ajout, inspection et retrait de Scripting avec restauration exacte des références.
3. Création et écriture de module/classe, accents français, refus d'une écriture avec hash périmé, export `.bas`/`.cls`, suppression et réimport avec gardes de versions.
4. Création d'un UserForm, ajout et dimensionnement d'un label, légende, police Arial 12 gras et relecture native.
5. Disponibilité de sauvegarde VBAi et refus sans mutation quand l'adaptateur est indisponible.
6. Navigation dans le code et compilation native du projet, avec capture des diagnostics et confirmation du mode conception.
7. Sauvegarde par l'API de l'application, réouverture puis relecture du code, des accents et du label avec sa police.

**La sauvegarde par l'API native du helper ne qualifie pas `save_host_document`.** Les refus répertoriés restent des lacunes ou prérequis, même lorsque le parcours documentaire réussit. Ces scénarios ne qualifient pas l'exécution de macros, les breakpoints/pas à pas, la signature, tous les contrôles, le rendu complet Monaco/thème, ni les fournisseurs LLM dans chaque hôte.

## Particularités des helpers

- Word peut démarrer une nouvelle instance même lorsque des instances précédentes sont présentes ; le test exige un nouveau PID avant toute mutation et ferme uniquement cette instance. `Normal` est exclu de la résolution du projet jetable.
- Les déclarations créées par l'hôte sont conservées, notamment `Option Compare Database` dans Access ; le code de test n'insère pas de déclaration après une fonction.
- Access peut afficher « Enregistrer sous » pour un module. UI Automation accepte uniquement le dialogue du PID possédé et les trois noms de composants du test. Aucune coordonnée ni raccourci global n'est utilisé.
- Access et Publisher sont fermés puis relancés pour vérifier la persistance. [Microsoft décrit la contrainte de nouvelle instance pour `Publisher.Application.Open`](https://learn.microsoft.com/en-us/office/vba/api/publisher.application.open). L'avertissement Publisher est traité par **Désactiver les macros**, sans modifier la politique de sécurité.
- Outlook ouvre uniquement un inspecteur non enregistré si un profil classique est déjà configuré, lit les métadonnées VBE puis le ferme avec `olDiscard`. Il ne lit ni n'envoie de courrier et ne modifie pas `VbaProject.OTM`.

## Relancer et lire les preuves

```powershell
# Compiler les tests contre la DLL actuellement installée, sans remplacer une DLL chargée.
dotnet build tests/VBAi.Tests/VBAi.Tests.csproj --no-restore -p:BuildProjectReferences=false
$env:VBAi_RUN_OFFICE_TESTS = '1'
$env:VBAi_RUN_OUTLOOK_TESTS = '1'
$env:VBAi_OFFICE_RESULTS = "$PWD/artifacts/office-hosts/qualification"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj --no-build --no-restore --filter 'TestCategory=Office' --logger 'trx;LogFileName=office-hosts.trx' --results-directory artifacts/office-hosts/vstest
```

Le rapport accepté est `artifacts/office-hosts/vstest/office-hosts-complete.trx`. Les fichiers `qualification.json`, exports et documents jetables sont conservés sous `artifacts/office-hosts/qualification/<hôte>/<identifiant>/`. Les rapports précédents restent présents pour le diagnostic ; leur date et leur résultat doivent être vérifiés avant de les utiliser.

Ces tests mesurent le comportement du complément chargé dans Office. Coverlet dans VSTest ne mesure pas les lignes exécutées dans les processus Office ; ce lot n'annonce donc aucun nouveau pourcentage global de couverture.
