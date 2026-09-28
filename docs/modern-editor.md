# Éditeur de code moderne intégré au VBE

## Objectif et périmètre accepté

Le 28 septembre 2026, le développement par étapes d’un éditeur moderne a été accepté. Il doit rester intégré au VBE, modifier les modules du projet ouvert et conserver la compilation et le débogage internes. Les numéros de ligne sont une marge visuelle : aucune numérotation n’est ajoutée au source VBA.

Fonctions visées : thème sombre, numéros de ligne, guides d’indentation, ligne active, coloration VBA, repliage des procédures et blocs, navigation par définition, puis marge de points d’arrêt et instruction courante.

## Jalons

1. Consolider le thème natif : cartographier les contrôles, gérer les fenêtres détruites/recréées et assurer le retrait du thème. La cartographie Excel est réalisée ; les menus Office, bandeaux et grilles VBA restent partiellement clairs.
2. Prototyper un éditeur ancrable, avec lecture et écriture explicites dans un module d’une macro jetable. Le choix du composant graphique et sa compatibilité .NET Framework 4.8 restent à vérifier.
3. Vérifier la synchronisation dans Excel, notamment les modifications concurrentes, erreurs et changements de mode. Faire valider ce jalon avant de généraliser son usage.
4. Ajouter les numéros, guides et coloration, puis le repliage VBA.
5. Intégrer les points d’arrêt et l’état du débogueur via les commandes internes existantes, sans raccourcis clavier.
6. Qualifier sous SOLIDWORKS 2020.

## Contrat de synchronisation du prototype

- Le document ouvert identifie le projet et le composant COM ; le nom seul ne suffit pas à identifier un module.
- Une copie du texte lu est conservée comme référence de synchronisation.
- Avant toute écriture, relire le module : si son texte diffère de cette référence, signaler un conflit et conserver les deux versions.
- Pour le premier prototype, utiliser une commande explicite d’application. Ne pas écrire à chaque frappe.
- Relire et vérifier le texte après écriture ; tenir compte des ajustements de casse et de format effectués par le VBE.
- Une édition de `CodeModule` n’est pas une transaction native. En cas d’échec partiel, conserver le tampon et la copie initiale ; ne pas annoncer un succès ou une restauration sans relecture.
- La première version autorise l’écriture uniquement en mode conception. En exécution ou en arrêt de débogage, conserver le tampon en lecture seule.
- Appliquer et vérifier les changements avant de lancer compilation ou exécution depuis le nouvel éditeur ; bloquer ce lancement en cas de conflit.
- L’écriture du module en mémoire et la sauvegarde du fichier hôte sont deux opérations distinctes. Le prototype ne doit pas annoncer une sauvegarde disque après la seule modification du module.
- L’annulation du nouvel éditeur porte sur son tampon. La coordination avec l’historique natif doit être étudiée avant toute promesse d’annulation unifiée.

## Contrôles du premier prototype

Sur un classeur jetable : ouvrir un module, modifier une procédure, appliquer puis relire ; modifier le même module depuis le VBE et vérifier le refus d’écrasement ; fermer un composant ou un projet avec un tampon modifié ; vérifier le refus d’écriture en débogage ; fermer Excel normalement.

L’intégration dans une fenêtre d’outil est documentée par Microsoft. Le remplacement transparent d’une fenêtre de document native ne l’est pas : le premier prototype sera donc ancrable et conservera l’accès au code natif.

## Références

- [CodeModule](https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.codemodule?view=office-pia)
- [Modèle des compléments VBA](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model)
- [CreateToolWindow](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtoolwindow-method)

## État

Le thème natif expérimental existe. Le nouvel éditeur et son mécanisme de synchronisation ne sont pas encore implémentés. Ce document fixe le périmètre de leur premier prototype.
