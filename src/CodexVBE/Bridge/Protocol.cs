namespace CodexVBE
{
    /// <summary>Arguments sérialisés en JSON pour une commande du pont entre le complément et son client.</summary>
    public sealed class Request
    {
        /// <summary>Nom de la commande demandée au pont.</summary>
        /// <value>Nom de la commande demandée au pont.</value>
        public string Command { get; set; }
        /// <summary>Nom du projet VBA visé par la commande.</summary>
        /// <value>Nom du projet VBA visé par la commande.</value>
        public string Project { get; set; }
        /// <summary>Nom du composant de code visé.</summary>
        /// <value>Nom du composant de code visé.</value>
        public string Module { get; set; }
        /// <summary>Première ligne concernée, selon la commande.</summary>
        /// <value>Première ligne concernée, selon la commande.</value>
        public int StartLine { get; set; }
        /// <summary>Première colonne concernée, selon la commande.</summary>
        /// <value>Première colonne concernée, selon la commande.</value>
        public int StartColumn { get; set; }
        /// <summary>Dernière ligne concernée, selon la commande.</summary>
        /// <value>Dernière ligne concernée, selon la commande.</value>
        public int EndLine { get; set; }
        /// <summary>Dernière colonne concernée, selon la commande.</summary>
        /// <value>Dernière colonne concernée, selon la commande.</value>
        public int EndColumn { get; set; }
        /// <summary>Nombre d’éléments demandé ou attendu.</summary>
        /// <value>Nombre d’éléments demandé ou attendu.</value>
        public int Count { get; set; }
        /// <summary>Empreinte attendue qui protège une modification contre un état périmé.</summary>
        /// <value>Empreinte attendue qui protège une modification contre un état périmé.</value>
        public string ExpectedSha256 { get; set; }
        /// <summary>Texte transmis ou attendu par la commande.</summary>
        /// <value>Texte transmis ou attendu par la commande.</value>
        public string Text { get; set; }
        /// <summary>Liste de chaînes transmise à la commande.</summary>
        /// <value>Liste de chaînes transmise à la commande.</value>
        public string[] Items { get; set; }
        /// <summary>Texte recherché ou filtre de la commande.</summary>
        /// <value>Texte recherché ou filtre de la commande.</value>
        public string Query { get; set; }
        /// <summary>Action demandée sur un élément.</summary>
        /// <value>Action demandée sur un élément.</value>
        public string Action { get; set; }
        /// <summary>Identifiant numérique du contrôle ciblé.</summary>
        /// <value>Identifiant numérique du contrôle ciblé.</value>
        public int ControlId { get; set; }
        /// <summary>Légende du contrôle à retrouver.</summary>
        /// <value>Légende du contrôle à retrouver.</value>
        public string ControlCaption { get; set; }
        /// <summary>Légende de la fenêtre VBE à retrouver.</summary>
        /// <value>Légende de la fenêtre VBE à retrouver.</value>
        public string WindowCaption { get; set; }
        /// <summary>Type numérique de la fenêtre VBE visée.</summary>
        /// <value>Type numérique de la fenêtre VBE visée.</value>
        public int WindowType { get; set; }
        /// <summary>Mode d’exécution attendu avant l’opération de débogage.</summary>
        /// <value>Mode d’exécution attendu avant l’opération de débogage.</value>
        public int ExpectedMode { get; set; }
        /// <summary>Nom du UserForm ciblé.</summary>
        /// <value>Nom du UserForm ciblé.</value>
        public string Form { get; set; }
        /// <summary>Nom du contrôle ciblé.</summary>
        /// <value>Nom du contrôle ciblé.</value>
        public string Control { get; set; }
        /// <summary>Type de contrôle à créer ou à vérifier.</summary>
        /// <value>Type de contrôle à créer ou à vérifier.</value>
        public string ControlType { get; set; }
        /// <summary>Position horizontale du contrôle.</summary>
        /// <value>Position horizontale du contrôle.</value>
        public double Left { get; set; }
        /// <summary>Position verticale du contrôle.</summary>
        /// <value>Position verticale du contrôle.</value>
        public double Top { get; set; }
        /// <summary>Largeur du contrôle.</summary>
        /// <value>Largeur du contrôle.</value>
        public double Width { get; set; }
        /// <summary>Hauteur du contrôle.</summary>
        /// <value>Hauteur du contrôle.</value>
        public double Height { get; set; }
        /// <summary>Légende à attribuer ou à comparer.</summary>
        /// <value>Légende à attribuer ou à comparer.</value>
        public string Caption { get; set; }
        /// <summary>Version attendue de l’arbre de contrôles du formulaire.</summary>
        /// <value>Version attendue de l’arbre de contrôles du formulaire.</value>
        public string ExpectedFormVersion { get; set; }
        /// <summary>Nom de la propriété ciblée.</summary>
        /// <value>Nom de la propriété ciblée.</value>
        public string Property { get; set; }
        /// <summary>Valeur à lire ou à attribuer à la propriété.</summary>
        /// <value>Valeur à lire ou à attribuer à la propriété.</value>
        public object Value { get; set; }
        /// <summary>Chemin associé à la commande, au fichier ou au composant.</summary>
        /// <value>Chemin associé à la commande, au fichier ou au composant.</value>
        public string Path { get; set; }
        /// <summary>Encodage demandé pour le texte source.</summary>
        /// <value>Encodage demandé pour le texte source.</value>
        public string SourceEncoding { get; set; }
        /// <summary>Identifiant GUID de la référence.</summary>
        /// <value>Identifiant GUID de la référence.</value>
        public string Guid { get; set; }
        /// <summary>Version majeure de la référence.</summary>
        /// <value>Version majeure de la référence.</value>
        public int Major { get; set; }
        /// <summary>Version mineure de la référence.</summary>
        /// <value>Version mineure de la référence.</value>
        public int Minor { get; set; }
        /// <summary>Position de départ dans une liste paginée.</summary>
        /// <value>Position de départ dans une liste paginée.</value>
        public int Offset { get; set; }
        /// <summary>Nombre maximal d’éléments à renvoyer.</summary>
        /// <value>Nombre maximal d’éléments à renvoyer.</value>
        public int Limit { get; set; }
        /// <summary>Index de ligne visé, ou null si la commande ne précise pas de ligne.</summary>
        /// <value>Index de ligne visé, ou null si la commande ne précise pas de ligne.</value>
        public int? RowIndex { get; set; }
        /// <summary>Version attendue de la liste avant sa modification.</summary>
        /// <value>Version attendue de la liste avant sa modification.</value>
        public string ExpectedListVersion { get; set; }
        /// <summary>Index du type de référence dans la liste du projet.</summary>
        /// <value>Index du type de référence dans la liste du projet.</value>
        public int TypeIndex { get; set; }
        /// <summary>Identité du type de référence à retrouver.</summary>
        /// <value>Identité du type de référence à retrouver.</value>
        public string TypeIdentity { get; set; }
        /// <summary>Version attendue des références avant leur modification.</summary>
        /// <value>Version attendue des références avant leur modification.</value>
        public string ExpectedReferencesVersion { get; set; }
        /// <summary>Chemin du conteneur parent dans l’arbre du formulaire.</summary>
        /// <value>Chemin du conteneur parent dans l’arbre du formulaire.</value>
        public string ParentPath { get; set; }
        /// <summary>Version attendue de l’arbre de contrôles avant sa modification.</summary>
        /// <value>Version attendue de l’arbre de contrôles avant sa modification.</value>
        public string ExpectedTreeVersion { get; set; }
        /// <summary>Chemin du contrôle dans l’arbre du formulaire.</summary>
        /// <value>Chemin du contrôle dans l’arbre du formulaire.</value>
        public string ControlPath { get; set; }
        /// <summary>Version attendue du projet avant sa modification.</summary>
        /// <value>Version attendue du projet avant sa modification.</value>
        public string ExpectedProjectVersion { get; set; }
        /// <summary>Chemin attendu du document hôte.</summary>
        /// <value>Chemin attendu du document hôte.</value>
        public string ExpectedHostPath { get; set; }
        /// <summary>Empreinte du certificat de signature à utiliser.</summary>
        /// <value>Empreinte du certificat de signature à utiliser.</value>
        public string CertificateThumbprint { get; set; }
        /// <summary>Version attendue des composants du projet.</summary>
        /// <value>Version attendue des composants du projet.</value>
        public string ExpectedComponentVersion { get; set; }
        /// <summary>Nouveau nom à attribuer à l’objet ciblé.</summary>
        /// <value>Nouveau nom à attribuer à l’objet ciblé.</value>
        public string NewName { get; set; }
        /// <summary>Nom de la procédure concernée.</summary>
        /// <value>Nom de la procédure concernée.</value>
        public string Procedure { get; set; }
        /// <summary>Nom de l’événement concerné.</summary>
        /// <value>Nom de l’événement concerné.</value>
        public string EventName { get; set; }
        /// <summary>Expression VBA à évaluer.</summary>
        /// <value>Expression VBA à évaluer.</value>
        public string Expression { get; set; }
        /// <summary>Nouvelle expression VBA à appliquer.</summary>
        /// <value>Nouvelle expression VBA à appliquer.</value>
        public string NewExpression { get; set; }
        /// <summary>Contexte textuel transmis à la commande.</summary>
        /// <value>Contexte textuel transmis à la commande.</value>
        public string Context { get; set; }
        /// <summary>Type de l’expression de surveillance.</summary>
        /// <value>Type de l’expression de surveillance.</value>
        public string WatchType { get; set; }
        /// <summary>Diagnostic à transmettre ou à afficher.</summary>
        /// <value>Diagnostic à transmettre ou à afficher.</value>
        public string Diagnostic { get; set; }
        /// <summary>Nom ou identifiant du bouton visé.</summary>
        /// <value>Nom ou identifiant du bouton visé.</value>
        public string Button { get; set; }
        /// <summary>Nom ou identifiant du volet visé.</summary>
        /// <value>Nom ou identifiant du volet visé.</value>
        public string Pane { get; set; }
        /// <summary>Segments ordonnés d’un chemin dans l’arbre du formulaire.</summary>
        /// <value>Segments ordonnés d’un chemin dans l’arbre du formulaire.</value>
        public string[] PathSegments { get; set; }
        /// <summary>Nom de l’objet visé.</summary>
        /// <value>Nom de l’objet visé.</value>
        public string ObjectName { get; set; }
        /// <summary>Code numérique du type de procédure VBA.</summary>
        /// <value>Code numérique du type de procédure VBA.</value>
        public int ProcKind { get; set; }
        /// <summary>Position à laquelle insérer l’élément, si elle est précisée.</summary>
        /// <value>Position à laquelle insérer l’élément, si elle est précisée.</value>
        public int? InsertIndex { get; set; }
        /// <summary>Ordre d’empilement demandé pour le contrôle.</summary>
        /// <value>Ordre d’empilement demandé pour le contrôle.</value>
        public int ZPosition { get; set; }
        /// <summary>Indique si la recherche doit correspondre à un mot entier.</summary>
        /// <value>Indique si la recherche doit correspondre à un mot entier.</value>
        public bool WholeWord { get; set; }
        /// <summary>Indique si la recherche respecte la casse.</summary>
        /// <value>Indique si la recherche respecte la casse.</value>
        public bool MatchCase { get; set; }
        /// <summary>Indique si la requête est interprétée comme un motif.</summary>
        /// <value>Indique si la requête est interprétée comme un motif.</value>
        public bool PatternSearch { get; set; }
        /// <summary>Indique si la capture de débogage inclut la pile d’appels.</summary>
        /// <value>Indique si la capture de débogage inclut la pile d’appels.</value>
        public bool IncludeCallStack { get; set; }
        /// <summary>Nom de la police à appliquer.</summary>
        /// <value>Nom de la police à appliquer.</value>
        public string FontName { get; set; }
        /// <summary>Taille de police à appliquer.</summary>
        /// <value>Taille de police à appliquer.</value>
        public double FontSize { get; set; }
        /// <summary>Indique si le texte doit être en gras.</summary>
        /// <value>Indique si le texte doit être en gras.</value>
        public bool FontBold { get; set; }
    }

    /// <summary>Résultat sérialisable renvoyé par le pont après traitement d’une commande.</summary>
    public sealed class Response
    {
        /// <summary>Indique si le traitement de la commande a réussi.</summary>
        /// <value>Vrai en cas de réussite.</value>
        public bool Ok { get; set; }
        /// <summary>Détail d’erreur renvoyé quand le traitement échoue.</summary>
        /// <value>Message d’erreur, ou null en cas de réussite.</value>
        public string Error { get; set; }
        /// <summary>Charge utile du résultat de la commande.</summary>
        /// <value>Donnée propre à la commande, ou null.</value>
        public object Data { get; set; }

        /// <summary>Construit une réponse positive contenant les données du résultat.</summary>
        /// <param name="data">Donnée renvoyée par la commande, éventuellement null.</param>
        /// <returns>Réponse marquée comme réussie.</returns>
        public static Response Success(object data) { return new Response { Ok = true, Data = data }; }
        /// <summary>Construit une réponse négative contenant le motif de l’échec.</summary>
        /// <param name="error">Motif destiné au client.</param>
        /// <returns>Réponse marquée comme échouée.</returns>
        public static Response Failure(string error) { return new Response { Ok = false, Error = error }; }
    }
}
