# Attributs multilignes : remplacement contrôlé autorisé

La vérification Excel montre que CodeModule.AddFromFile ignore VB_Description après une déclaration continuée, contrairement à VBComponents.Import. Un remplacement préparé est utilisé pour les composants indépendants. Les modules de document restent protégés.

## Remplacement

1. Exporter le composant original avec son éventuel fichier FRX dans un répertoire de récupération.
2. Préparer le fichier modifié en conservant les attributs et les données du Designer.
3. Importer un candidat temporaire sous un nom distinct ; vérifier le code et les attributs par un nouvel export.
4. Relire le composant original et vérifier son identité, son texte, son mode et la protection du projet. Recontrôler aussi son snapshot Designer juste avant retrait pour préserver une modification concurrente.
5. Donner un nom de retrait temporaire à l'original, attribuer le nom final au candidat, revalider le candidat et le texte original, puis retirer l'original et reconnecter l'onglet Monaco à sa nouvelle identité COM. Cette séquence évite le verrouillage du nom par le Designer UserForm encore chargé.
6. En cas d'échec, restaurer l'export original ; conserver les fichiers et signaler leur chemin si la restauration échoue.

Portée : modules standards, classes et UserForms. Les modules de document (feuille/classeur) restent exclus car leur identité est détenue par Excel.

Risques : nouvelle identité COM, invalidation des références externes au composant, perte possible des points d'arrêt et de l'historique Undo. Les macros ne sont pas exécutées. La validation utilise exclusivement un classeur jetable.

La revue automatique a initialement refusé le remplacement, puis l'utilisateur l'a explicitement autorisé. Les opérations suivantes utilisent cette autorisation, sans nouvelle demande pour chaque édition.

Le candidat reçoit un nom unique avant import ; pour un UserForm, ce nom est adapté dans l'en-tête Designer et VB_Name. Le FRX original est réutilisé tel quel à l'import. Le réexport OLE contient des horodatages et du remplissage non déterministes : la vérification porte sur la hiérarchie, les propriétés persistantes lisibles, les images et polices exposées par le snapshot Designer, avant retrait puis après renommage. Les contrôles tiers et propriétés opaques non vérifiables bloquent le remplacement. Excel ajoute une ligne vide initiale à chaque import UserForm. Seul ce préfixe vide supplémentaire est retiré avant les déclarations, puis le code et les attributs sont relus, afin de restaurer exactement le texte et de ne pas accumuler des lignes vides.

La qualification comprend un échec injecté après retrait pour un module standard et un UserForm, ainsi qu'une modification concurrente du bouton du Designer pendant la préparation. Le texte, les attributs, les contrôles et le nombre de composants sont relus après restauration/refus.
