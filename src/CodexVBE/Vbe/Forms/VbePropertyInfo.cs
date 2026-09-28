using System.Collections.Generic;

namespace CodexVBE
{
    /// <summary>Description sérialisable d’une propriété de formulaire ou d’un membre imbriqué.</summary>
    internal sealed class VbePropertyInfo
    {
        /// <summary>Nom COM de la propriété.</summary>
        /// <value>Nom de propriété fourni par COM.</value>
        public string Name { get; set; }
        /// <summary>Nom du type observé pour la valeur.</summary>
        /// <value>Nom du type descriptif sérialisable.</value>
        public string Type { get; set; }
        /// <summary>Catégorie de valeur ou de membre décrite.</summary>
        /// <value>Type logique de propriété ou de membre.</value>
        public string Kind { get; set; }
        /// <summary>État en lecture seule déclaré par le descripteur COM, s’il est disponible.</summary>
        /// <value>true/false selon le descripteur, ou null si inconnu.</value>
        public bool? ReadOnly { get; set; }
        // Names from the live enum type, independent of its generated COM type name.
        // This is a read-only catalog, not proof that the native setter works.
        /// <summary>Noms de valeurs enum détectés, sans preuve de succès d’écriture.</summary>
        /// <value>Valeurs enum connues, ou null si la propriété n’est pas une enum.</value>
        public string[] AllowedValues { get; set; }
        // COM descriptors can advertise a setter that fails at invocation
        // (observed for Label.Cancel). This is metadata, not runtime proof.
        /// <summary>État descriptif de la prise en charge du setter.</summary>
        private string setterStatus;
        /// <summary>État du setter déclaré ou candidat d’écriture non vérifié.</summary>
        /// <value>Statut dérivé du descripteur, sauf si une valeur explicite est enregistrée.</value>
        public string SetterStatus
        {
            get { return setterStatus ?? (ReadOnly == true ? "DescriptorReadOnly" :
                ReadOnly == false ? "DescriptorCandidateUnverified" : "Unknown"); }
            set { setterStatus = value; }
        }
        /// <summary>Valeur lue, si elle a pu être obtenue.</summary>
        /// <value>Valeur sérialisée, ou null si la lecture a échoué.</value>
        public object Value { get; set; }
        /// <summary>Représentation destinée à l’affichage.</summary>
        /// <value>Texte compact et adapté à l’inspection.</value>
        public string Display { get; set; }
        /// <summary>Empreinte des valeurs non exposées directement.</summary>
        /// <value>Empreinte utilisée pour comparer une valeur non affichée.</value>
        public string Digest { get; set; }
        /// <summary>Erreur rencontrée pendant la lecture de cette valeur.</summary>
        /// <value>Message d’erreur associé à la lecture.</value>
        public string Error { get; set; }
        /// <summary>Nombre d’indices attendus par la propriété COM.</summary>
        /// <value>Nombre d’arguments d’index COM.</value>
        public int NumIndices { get; set; }
        /// <summary>Descriptions des membres imbriqués accessibles.</summary>
        /// <value>Propriétés décrites sur l’objet imbriqué.</value>
        public List<VbePropertyInfo> Members { get; set; }
    }
}
