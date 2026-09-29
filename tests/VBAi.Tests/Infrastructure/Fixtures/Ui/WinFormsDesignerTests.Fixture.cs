namespace VBAi.Tests.Unit
{
    using System;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit une recherche vérifiable des contrôles créés par les fenêtres Designer.</summary>
    public sealed partial class WinFormsDesignerTests
    {
        /// <summary>Recherche un unique contrôle par nom dans l’arbre visuel.</summary>
        /// <typeparam name="T">Type de contrôle attendu.</typeparam>
        /// <param name="root">Racine à parcourir.</param>
        /// <param name="name">Nom du contrôle recherché.</param>
        /// <returns>Contrôle trouvé, converti en <typeparamref name="T"/>.</returns>
        private static T Find<T>(Control root, string name)
            where T : Control
        {
            var matches = root.Controls.Find(name, true);
            Assert.AreEqual(1, matches.Length, "Missing or duplicated designer control: " + name);
            return (T)matches[0];
        }
    }
}
