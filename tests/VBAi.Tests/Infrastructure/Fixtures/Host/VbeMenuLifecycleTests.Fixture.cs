namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Construit un hôte et des menus simulés pour tester le cycle de vie des commandes VBE.</summary>
    public sealed partial class VbeMenuLifecycleTests
    {
        /// <summary>Crée une barre principale et une barre de fenêtre de code contenant les menus attendus.</summary>
        /// <returns>Hôte simulé avec les barres et commandes View et Tools.</returns>
        private static FakeHost Host()
        {
            var view = new FakeButton
            {
                Caption = "&View"
            };
            var tools = new FakeButton
            {
                Caption = "&Tools"
            };
            var main = new FakeBar
            {
                Type = 1,
                Controls = new FakeControls()
            };
            main.Controls.Items.Add(view);
            main.Controls.Items.Add(tools);
            return new FakeHost
            {
                CommandBars = new[]
                {
                    main,
                    new FakeBar
                    {
                        Type = 0,
                        Name = "Code Window",
                        Controls = new FakeControls()
                    }
                }
            };
        }

        /// <summary>Hôte simulé exposant la collection de barres de commandes.</summary>
        public sealed class FakeHost
        {
            /// <summary>Barres disponibles dans l’hôte.</summary>
            /// <value>Tableau des barres de commandes.</value>
            public FakeBar[] CommandBars { get; set; }
        }

        /// <summary>Barre simulée avec son type, son nom et ses contrôles.</summary>
        public sealed class FakeBar
        {
            /// <summary>Type numérique de la barre.</summary>
            /// <value>Type lu pendant l’installation des commandes.</value>
            public int Type { get; set; }
            /// <summary>Nom de la barre lorsqu’elle est nommée.</summary>
            /// <value>Nom fourni au cycle de vie du menu.</value>
            public string Name { get; set; }
            /// <summary>Collection de contrôles appartenant à la barre.</summary>
            /// <value>Contrôles modifiables de la barre.</value>
            public FakeControls Controls { get; set; }
        }

        /// <summary>Commande simulée avec ses métadonnées et son historique de suppression.</summary>
        public sealed class FakeButton
        {
            /// <summary>Libellé affiché de la commande.</summary>
            /// <value>Texte du bouton.</value>
            public string Caption { get; set; }
            /// <summary>Marqueur d’identification de la commande.</summary>
            /// <value>Valeur utilisée pour retrouver le bouton du complément.</value>
            public string Tag { get; set; }
            /// <summary>Texte d’aide affiché au survol.</summary>
            /// <value>Info-bulle du bouton.</value>
            public string TooltipText { get; set; }
            public object Picture { get; set; }
            public object Mask { get; set; }
            public int Style { get; set; }
            public bool RejectDelete { get; set; }
            /// <summary>Contrôles enfants du bouton.</summary>
            /// <value>Collection initialisée lors de la création du bouton.</value>
            public FakeControls Controls { get; } = new FakeControls();
            /// <summary>Nombre d’appels à <see cref="Delete"/>.</summary>
            /// <value>Compteur de suppressions.</value>
            public int DeleteCount { get; private set; }

            /// <summary>Incrémente le compteur de suppressions du bouton.</summary>
            public void Delete()
            {
                if (RejectDelete) throw new InvalidOperationException("Delete rejected");
                DeleteCount++;
            }
        }

        /// <summary>Collection énumérable qui simule l’ajout de commandes et les échecs d’insertion.</summary>
        public sealed class FakeControls : IEnumerable<FakeButton>
        {
            /// <summary>Boutons actuellement présents.</summary>
            /// <value>Liste modifiable des commandes.</value>
            public List<FakeButton> Items { get; } = new List<FakeButton>();
            /// <summary>Nombre d’éléments à partir duquel l’ajout doit échouer.</summary>
            /// <value>Limite par défaut à <see cref="int.MaxValue"/>.</value>
            public int FailAtCount { get; set; } = int.MaxValue;

            /// <summary>Ajoute un bouton, sauf si le seuil d’échec configuré est atteint.</summary>
            /// <param name="type">Type de contrôle demandé.</param>
            /// <param name="id">Identifiant de commande facultatif.</param>
            /// <param name="parameter">Paramètre de création facultatif.</param>
            /// <param name="before">Contrôle avant lequel insérer le bouton.</param>
            /// <param name="temporary">Indique si le bouton est temporaire.</param>
            /// <returns>Bouton créé et ajouté à la collection.</returns>
            /// <exception cref="InvalidOperationException">Le nombre actuel d’éléments a atteint <see cref="FailAtCount"/>.</exception>
            public FakeButton Add(int type, object id, object parameter, object before, bool temporary)
            {
                if (Items.Count >= FailAtCount)
                    throw new InvalidOperationException("Add failed");
                var button = new FakeButton();
                Items.Add(button);
                return button;
            }

            /// <summary>Retourne un énumérateur générique sur les boutons.</summary>
            /// <returns>Énumérateur de la liste des boutons.</returns>
            public IEnumerator<FakeButton> GetEnumerator()
            {
                return Items.GetEnumerator();
            }

            /// <summary>Retourne l’énumérateur non générique de la collection.</summary>
            /// <returns>Énumérateur des boutons.</returns>
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
