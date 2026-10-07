namespace VBAi
{

    /// <summary>Décrit un modèle proposé par un fournisseur LLM et ses réglages de niveau.</summary>
    internal sealed class LlmModelOption
    {

        /// <summary>Identifiant du modèle transmis au fournisseur.</summary>
        /// <value>Identifiant utilisé dans les requêtes du fournisseur.</value>
        public string Id { get; private set; }

        /// <summary>Libellé présenté dans la liste des modèles.</summary>
        /// <value>Libellé fourni, ou identifiant si le libellé est vide.</value>
        public string Label { get; private set; }

        /// <summary>Indique si le fournisseur désigne ce modèle comme choix par défaut.</summary>
        /// <value>Vrai si ce modèle est le choix par défaut.</value>
        public bool IsDefault { get; private set; }

        /// <summary>Niveau sélectionné par défaut pour le modèle.</summary>
        /// <value>Identifiant du niveau par défaut, ou null.</value>
        public string DefaultEffort { get; private set; }

        /// <summary>Niveaux proposés pour le modèle.</summary>
        /// <value>Options de niveau, ou tableau vide si aucune option est fournie.</value>
        public LlmEffortOption[] Efforts { get; private set; }

        /// <summary>Explicit provider capability declarations for this catalog entry.</summary>
        /// <value>Immutable declarations; unknown fields are represented by null.</value>
        public LlmModelCapabilities Capabilities { get; }

        /// <summary>Crée une option de modèle et normalise son libellé et sa liste de niveaux.</summary>
        /// <param name="id">Identifiant du modèle côté fournisseur.</param>
        /// <param name="label">Libellé affiché ; si vide, identifiant utilisé.</param>
        /// <param name="isDefault">Indique si le modèle est proposé par défaut.</param>
        /// <param name="defaultEffort">Identifiant du niveau par défaut, éventuellement null.</param>
        /// <param name="efforts">Niveaux pris en charge, éventuellement null.</param>
        /// <param name="capabilities">Explicit catalog capabilities, or null when no declaration is available.</param>
        public LlmModelOption(string id, string label, bool isDefault = false,
            string defaultEffort = null, LlmEffortOption[] efforts = null, LlmModelCapabilities capabilities = null)
        {
            Id = id;
            Label = string.IsNullOrWhiteSpace(label) ? id : label;
            IsDefault = isDefault;
            DefaultEffort = defaultEffort;
            Efforts = efforts ?? new LlmEffortOption[0];
            Capabilities = capabilities ?? LlmModelCapabilities.Unknown;
        }

        /// <summary>Formate le libellé et l’identifiant pour l’affichage d’une option.</summary>
        /// <returns>Identifiant seul si libellé et identifiant sont égaux ; sinon libellé suivi de l’identifiant.</returns>
        public override string ToString() { return Label == Id ? Id : Label + " (" + Id + ")"; }
    }

    /// <summary>Décrit un niveau utilisable lors d’une requête au modèle.</summary>
    internal sealed class LlmEffortOption
    {

        /// <summary>Identifiant reconnu par le fournisseur.</summary>
        /// <value>Identifiant envoyé dans les requêtes.</value>
        public string Id { get; private set; }

        /// <summary>Texte qui explique le niveau dans l’interface.</summary>
        /// <value>Description de l’option.</value>
        public string Description { get; private set; }

        /// <summary>Crée un niveau avec son identifiant fournisseur et sa description.</summary>
        /// <param name="id">Identifiant utilisé par le fournisseur.</param>
        /// <param name="description">Explication affichée.</param>
        public LlmEffortOption(string id, string description) { Id = id; Description = description; }

        /// <summary>Retourne l’identifiant pour l’affichage.</summary>
        /// <returns>Identifiant du niveau.</returns>
        public override string ToString() { return Id; }
    }
}
