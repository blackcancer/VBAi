using System;
using System.Drawing;
using System.Reflection;
using VBAi;
namespace VBAi.Tests.Infrastructure
{
    /// <summary>Sauvegarde les délégués et préférences de thème puis les restaure à la fin du test.</summary>
    internal sealed class ThemeScope : IDisposable
    {
        /// <summary>Choix de thème initial.</summary>
        private readonly ThemeChoice choice = UiTheme.Choice;
        /// <summary>Chemin du fichier de thème initial.</summary>
        private readonly string file = UiTheme.FileName;
        /// <summary>Détecteur de contraste initial.</summary>
        private readonly Func<bool> contrast = UiTheme.HighContrast;
        /// <summary>Lecteur de couleur de fenêtre initial.</summary>
        private readonly Func<Color> window = UiTheme.WindowColor;
        /// <summary>Lecteur de thème système initial.</summary>
        private readonly Func<object> system = UiTheme.ReadSystemTheme;
        /// <summary>Lecteur de préférences de thème initial.</summary>
        private readonly Func<string,string> read = UiTheme.ReadTheme;
        /// <summary>Écrivain de préférences de thème initial.</summary>
        private readonly Action<string,string> write = UiTheme.WriteTheme;
        /// <summary>Neutralise l’accès système et sélectionne un thème clair ou sombre pour le test.</summary>
        /// <param name="dark"><see langword="true"/> pour sélectionner le thème sombre.</param>
        internal ThemeScope(bool dark = false)
        {
            UiTheme.HighContrast = () => false;
            UiTheme.WriteTheme = (p,v) => { };
            SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
        }
        /// <summary>Change le choix de thème interne à l’interface.</summary>
        /// <param name="value">Choix de thème à appliquer.</param>
        internal static void SetChoice(ThemeChoice value) { typeof(UiTheme).GetProperty("Choice", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null,value); }
        /// <summary>Restaure toutes les préférences et fonctions de thème sauvegardées.</summary>
        public void Dispose() { SetChoice(choice); UiTheme.FileName=file; UiTheme.HighContrast=contrast; UiTheme.WindowColor=window; UiTheme.ReadSystemTheme=system; UiTheme.ReadTheme=read; UiTheme.WriteTheme=write; }
    }
    /// <summary>Contexte de licence qui indique que les composants s’exécutent dans le concepteur.</summary>
    internal sealed class DesignContext : System.ComponentModel.LicenseContext
    {
        /// <summary>Indique que le contexte est utilisé à la conception.</summary>
        /// <value>Retourne toujours <see cref="System.ComponentModel.LicenseUsageMode.Designtime"/>.</value>
        public override System.ComponentModel.LicenseUsageMode UsageMode { get { return System.ComponentModel.LicenseUsageMode.Designtime; } }
    }
    /// <summary>Accède aux champs et méthodes privés utilisés par les tests UI.</summary>
    internal static class UiInvoke
    {
        /// <summary>Lit un champ d’instance non public sur un objet de test.</summary>
        /// <typeparam name="T">Type attendu de la valeur du champ.</typeparam>
        /// <param name="target">Objet contenant le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        internal static T Field<T>(object target,string name) { return (T)target.GetType().GetField(name,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(target); }
        /// <summary>Appelle une méthode statique ou d’instance non publique par réflexion.</summary>
        /// <param name="type">Type qui déclare la méthode.</param>
        /// <param name="method">Nom de la méthode à appeler.</param>
        /// <param name="target">Instance cible, ou <see langword="null"/> pour une méthode statique.</param>
        /// <param name="args">Arguments de l’appel.</param>
        /// <returns>Valeur renvoyée par la méthode.</returns>
        internal static object Call(Type type, string method, object target, params object[] args)
        {
            return type.GetMethod(method, BindingFlags.Static | BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target,args);
        }
    }
}
