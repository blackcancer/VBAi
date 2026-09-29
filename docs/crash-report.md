# Rapports de problème et d’erreur

## Accès et envoi

**Outils → Signaler un problème** ouvre une fenêtre WinForms indépendante,
utilisable lorsque le chat est fermé. Le titre et les circonstances sont éditables ;
l’aperçu montre le rapport complet avant l’envoi. Copier, Enregistrer le rapport,
E-mail et Fermer restent disponibles sans authentification.

**Envoyer le rapport** utilise le compte GitHub sélectionné dans les paramètres
et ses identifiants Git Credential Manager, sans ouvrir une connexion interactive.
L’issue est créée dans `blackcancer/VBAi`, jamais dans le dépôt de la macro.
Ce dépôt était privé lors de la vérification du 28 septembre 2026 : le compte doit
avoir accès aux issues. Le lien vers l’issue créée apparaît après une réponse valide.

En cas d’absence d’identifiants ou de refus HTTP 4xx, VBAi utilise Outlook classique,
s’il dispose d’un profil et d’un compte, pour envoyer à **init-sys-rev@hotmail.com**.
L’envoi passe par le compte Outlook par défaut ; le succès signifie remis à Outlook,
pas réception garantie. VBAi ne quitte pas Outlook et ne crée pas de profil.
Le nouveau Outlook sans serveur COM ne permet pas cet envoi automatisé.

Si Outlook n’est pas disponible, l’application mail par défaut reçoit un brouillon
`mailto:`. Celui-ci contient le titre et le chemin de la copie locale : **joindre le
fichier enregistré puis envoyer le brouillon**. L’URI reste courte pour la compatibilité
avec les clients Windows. Le bouton Enregistrer recopie le chemin ; Copier fournit
le contenu complet. Aucun serveur SMTP ni mot de passe mail n’est demandé.

Un timeout, une connexion interrompue, un retour GitHub invalide, une erreur HTTP 5xx
ou un échec pendant `Outlook.Send` produit un état **envoi incertain**. Les boutons
d’envoi sont alors bloqués pour ce rapport : vérifier GitHub/Outlook avant de créer
un nouveau rapport. Aucun second envoi automatique ne masque un premier envoi possible.

## Contenu et persistance

Le rapport technique contient un identifiant, l’heure UTC, la version VBAi, le nom
du processus hôte, l’architecture, le CLR, la langue et le thème. Pour une erreur,
il inclut le type et les noms des méthodes de la pile, sans arguments ni fichiers source.
Les messages bruts d’exception, `Data`, journaux, réglages, clés, conversations et code
VBA ne sont pas joints. La description saisie est incluse volontairement et visible
dans l’aperçu ; elle peut donc contenir les informations que l’utilisateur y ajoute.

Une copie Markdown est enregistrée avant tout envoi dans
`%LocalAppData%\VBAi\CrashReports\<identifiant>.md`. Si cette sauvegarde échoue,
l’envoi est interrompu. Cette zone reste privée au profil Windows et ne crée aucun
fichier autour de la macro. Aucune suppression automatique des copies n’est effectuée.

## Capture automatique : périmètre exact

- Les erreurs inattendues attribuables à VBAi interceptées pendant sa connexion,
  l’ouverture du chat et les actions Paramètres/GitHub proposent le rapport ; les
  erreurs opérationnelles `ArgumentException`, `InvalidOperationException`, COM et
  annulations conservent leur traitement habituel.
- Les erreurs managées non gérées de l’AppDomain et les tâches non observées, lorsqu’une
  pile passe par l’assembly VBAi, sauvegardent un rapport et un marqueur `.pending`.
  Aucune fenêtre ni transmission réseau n’est tentée pendant une terminaison fatale.
- Au prochain chargement réussi du complément, un rapport en attente est proposé.
  Après fermeture du dialogue, son marqueur disparaît ; le Markdown reste conservé.
  Un seul rapport en attente est proposé par connexion.
- Les abonnements sont retirés à la déconnexion et une garde empêche la récursion.
  Aucun mode d’exception global de WinForms n’est modifié ; les événements étrangers
  au complément ne sont pas revendiqués ni marqués comme observés.

La capture ne couvre pas toutes les erreurs UI absorbées par l’hôte, les crashs
natifs VBE/Office/SOLIDWORKS, les accès mémoire invalides, un manque de mémoire sévère
ou une terminaison forcée. Ce mécanisme ne remplace pas un dump natif.

## Designer, langues et validation

`src/VBAi/Ui/CrashReportWindow.cs`, `.Designer.cs` et `.resx` constituent un
formulaire standard avec 20 contrôles fixes. Seules les données, les états de livraison
et les couleurs sont actualisés à l’exécution. Les textes sont traduits hors ligne
dans les 13 catalogues existants ; les champs techniques restent de gauche à droite.

Les transports des tests sont simulés : aucun rapport GitHub réel ni e-mail réel
n’a été envoyé. La compilation et les contrôles DesignSurface sont locaux ; le parcours
Outlook natif et l’ouverture/enregistrement manuel du source dans Visual Studio restent
à valider. Les captures autonomes n’utilisent ni VBE ni macro.

Validation : une suite complète de **1 150 tests unitaires** a réussi avant les
ajustements finaux de présentation et le test de refus sur échec de sauvegarde.
Les scénarios du rapport, des menus, du cycle de vie AddIn et de la localisation
ont ensuite été rejoués : **41 tests ciblés réussis**, puis **3 tests de livraison**
après le dernier ajustement du repli sur expiration GCM. **29 surfaces / 302 contrôles**
passent le chargement DesignSurface, les modifications de propriété et la
sérialisation/désérialisation. Aucune mesure de couverture instrumentée n’a été lancée.

Preuves locales : `artifacts/crash-report/tests/`, `artifacts/crash-report/designers/`
et `artifacts/crash-report/screens/`.
