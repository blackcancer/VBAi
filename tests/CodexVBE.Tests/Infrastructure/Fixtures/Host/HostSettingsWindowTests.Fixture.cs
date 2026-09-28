namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit des accès aux contrôles privés et des menus simulés pour les tests de fenêtre.</summary>
    public sealed partial class HostSettingsWindowTests
    {
        /// <summary>Options de réflexion pour les champs d’instance privés.</summary>
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        /// <summary>Lit un champ privé de la fenêtre de configuration.</summary>
        /// <typeparam name="T">Type attendu du champ.</typeparam>
        /// <param name="owner">Fenêtre contenant le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        private static T Field<T>(object owner, string name)
        {
            var field = owner.GetType().GetField(name, Private);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(owner);
        }

        /// <summary>Hôte minimal contenant les barres de menus simulées.</summary>
        public sealed class FakeMenus
        {
            /// <summary>Barres de commandes visibles de l’hôte simulé.</summary>
            /// <value>Liste modifiable des barres.</value>
            public List<FakeBar> CommandBars { get; } = new List<FakeBar>();
        }

        /// <summary>Barre simulée contenant les menus de l’hôte.</summary>
        public sealed class FakeBar
        {
            /// <summary>Type numérique de la barre.</summary>
            /// <value>Type utilisé par la reconnaissance des menus.</value>
            public int Type { get; set; }
            /// <summary>Menus présents dans la barre.</summary>
            /// <value>Liste de contrôles de menu.</value>
            public List<FakeMenu> Controls { get; set; }
        }

        /// <summary>Contrôle de menu simulé avec un libellé localisé.</summary>
        public sealed class FakeMenu
        {
            /// <summary>Libellé visible du menu.</summary>
            /// <value>Texte lu par les tests de reconnaissance de langue.</value>
            public string Caption { get; set; }
        }
    }
}
