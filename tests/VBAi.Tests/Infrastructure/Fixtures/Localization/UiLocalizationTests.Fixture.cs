namespace VBAi.Tests.Unit
{
    /// <summary>Fournit des hôtes simulés aux tests de localisation.</summary>
    public sealed partial class UiLocalizationTests
    {
        /// <summary>Crée un hôte dont une commande porte le libellé indiqué.</summary>
        /// <param name="caption">Libellé de la commande simulée.</param>
        /// <returns>Hôte de test contenant une barre et une commande.</returns>
        private static FakeVbe Host(string caption)
        {
            return new FakeVbe
            {
                CommandBars = new[]
                {
                    new FakeBar
                    {
                        Type = 1,
                        Controls = new[]
                        {
                            new FakeControl
                            {
                                Caption = caption
                            }
                        }
                    }
                }
            };
        }

        /// <summary>Représente l’hôte VBE minimal consommé par la détection de langue.</summary>
        public sealed class FakeVbe
        {
            /// <summary>Barres de commandes exposées par l’hôte simulé.</summary>
            /// <value>Tableau des barres disponibles.</value>
            public FakeBar[] CommandBars { get; set; }
        }

        /// <summary>Représente une barre de commandes minimale.</summary>
        public sealed class FakeBar
        {
            /// <summary>Type de barre de commandes.</summary>
            /// <value>Valeur numérique du type de barre.</value>
            public int Type { get; set; }
            /// <summary>Commandes contenues dans la barre.</summary>
            /// <value>Contrôles simulés de la barre.</value>
            public FakeControl[] Controls { get; set; }
        }

        /// <summary>Représente un contrôle avec son seul libellé utile à la localisation.</summary>
        public sealed class FakeControl
        {
            /// <summary>Libellé affiché du contrôle.</summary>
            /// <value>Texte utilisé pour reconnaître la langue du menu.</value>
            public string Caption { get; set; }
        }
    }
}
