using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Threading;
using CodexVBE;
namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Sauvegarde et restaure la culture UI du complément et les cultures du thread courant.</summary>
    internal sealed class LocalizationScope : IDisposable
    {
        /// <summary>Culture du complément avant le test.</summary>
        private readonly CultureInfo culture=UiText.Culture;
        /// <summary>Culture du thread avant le test.</summary>
        private readonly CultureInfo hostCulture=Thread.CurrentThread.CurrentCulture;
        /// <summary>Culture UI du thread avant le test.</summary>
        private readonly CultureInfo hostUiCulture=Thread.CurrentThread.CurrentUICulture;
        /// <summary>Sélectionne la culture UI utilisée pendant la portée.</summary>
        /// <param name="name">Nom de culture pris en charge, par défaut <c>en-US</c>.</param>
        internal LocalizationScope(string name="en-US") { Set(name); }
        /// <summary>Change la culture statique interne de l’interface.</summary>
        /// <param name="name">Nom de culture à appliquer.</param>
        internal static void Set(string name) { typeof(UiText).GetProperty("Culture",BindingFlags.Static|BindingFlags.NonPublic).SetValue(null,CultureInfo.GetCultureInfo(name)); }
        /// <summary>Restaure la culture de l’interface et les cultures du thread sauvegardées.</summary>
        public void Dispose() { Set(culture.Name); Thread.CurrentThread.CurrentCulture=hostCulture; Thread.CurrentThread.CurrentUICulture=hostUiCulture; }
    }
    /// <summary>Ressource de test qui ne fournit aucune traduction.</summary>
    internal sealed class EmptyResourceManager : ResourceManager
    {
        /// <summary>Retourne toujours <see langword="null"/> pour simuler une traduction absente.</summary>
        /// <param name="name">Clé de ressource demandée.</param>
        /// <param name="culture">Culture demandée.</param>
        /// <returns><see langword="null"/> quelle que soit la clé et la culture.</returns>
        public override string GetString(string name,CultureInfo culture) { return null; }
    }
}
