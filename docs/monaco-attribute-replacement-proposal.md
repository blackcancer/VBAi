# Attributs multilignes : opération proposée, en attente d'autorisation

La vérification Excel montre que CodeModule.AddFromFile ignore VB_Description après une déclaration continuée, contrairement à VBComponents.Import. La protection refuse actuellement cette édition avant toute écriture.

## Remplacement proposé

1. Exporter le composant original avec son éventuel fichier FRX dans un répertoire de récupération.
2. Préparer le fichier modifié en conservant les attributs et les données du Designer.
3. Importer un candidat temporaire sous un nom distinct ; vérifier le code et les attributs par un nouvel export.
4. Relire le composant original et vérifier son identité, son texte, son mode et la protection du projet.
5. Retirer l'original, renommer le candidat et reconnecter l'onglet Monaco à sa nouvelle identité COM.
6. En cas d'échec, restaurer l'export original ; conserver les fichiers et signaler leur chemin si la restauration échoue.

Portée proposée : modules standards, classes et UserForms. Les modules de document (feuille/classeur) restent exclus car leur identité est détenue par Excel.

Risques : nouvelle identité COM, invalidation des références externes au composant, perte possible des points d'arrêt et de l'historique Undo. Les macros ne sont pas exécutées. La validation commencerait uniquement dans un classeur jetable.

La revue automatique a refusé cette opération faute d'autorisation explicite du remplacement. Aucun chemin de remplacement n'a été installé ni exécuté dans le produit.
