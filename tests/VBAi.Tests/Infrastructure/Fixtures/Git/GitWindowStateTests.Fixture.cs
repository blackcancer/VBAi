namespace VBAi.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.Serialization;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit des accès réfléchis aux contrôles privés de la fenêtre Git pendant les tests.</summary>
    public sealed partial class GitWindowStateTests
    {
        /// <summary>Options de réflexion permettant d’accéder aux membres d’instance privés.</summary>
        private const BindingFlags InstancePrivate = BindingFlags.Instance | BindingFlags.NonPublic;
        /// <summary>Lit un champ privé de la fenêtre Git.</summary>
        /// <typeparam name="T">Type attendu du champ.</typeparam>
        /// <param name="window">Fenêtre contenant le champ.</param>
        /// <param name="name">Nom du champ privé.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        private static T Field<T>(GitWindow window, string name)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(window);
        }

        /// <summary>Écrit un champ privé de la fenêtre Git.</summary>
        /// <param name="window">Fenêtre contenant le champ.</param>
        /// <param name="name">Nom du champ privé.</param>
        /// <param name="value">Valeur à affecter.</param>
        private static void Set(GitWindow window, string name, object value)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            field.SetValue(window, value);
        }

        /// <summary>Appelle une méthode privée de la fenêtre Git.</summary>
        /// <param name="window">Fenêtre qui porte la méthode.</param>
        /// <param name="name">Nom de la méthode privée.</param>
        /// <param name="args">Arguments transmis à la méthode.</param>
        /// <returns>Valeur renvoyée par la méthode appelée.</returns>
        private static object Invoke(GitWindow window, string name, params object[] args)
        {
            var method = typeof(GitWindow).GetMethod(name, InstancePrivate);
            Assert.IsNotNull(method, name);
            if (name == "Perform" && args.Length == 1)
                args = new[]
                {
                    args[0],
                    (object)false
                };
            return method.Invoke(window, args);
        }

        /// <summary>Récupère la grille de différences interne à une vue de code.</summary>
        /// <param name="view">Vue contenant la grille privée.</param>
        /// <returns>Grille de différences de la vue.</returns>
        private static DataGridView DiffGrid(CodeDiffView view)
        {
            var field = typeof(CodeDiffView).GetField("grid", InstancePrivate);
            Assert.IsNotNull(field);
            return (DataGridView)field.GetValue(view);
        }
    }
}
