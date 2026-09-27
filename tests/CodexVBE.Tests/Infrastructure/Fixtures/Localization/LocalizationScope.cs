using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Threading;
using CodexVBE;
namespace CodexVBE.Tests.Infrastructure
{
    internal sealed class LocalizationScope : IDisposable
    {
        private readonly CultureInfo culture=UiText.Culture;
        private readonly CultureInfo hostCulture=Thread.CurrentThread.CurrentCulture;
        private readonly CultureInfo hostUiCulture=Thread.CurrentThread.CurrentUICulture;
        internal LocalizationScope(string name="en-US") { Set(name); }
        internal static void Set(string name) { typeof(UiText).GetProperty("Culture",BindingFlags.Static|BindingFlags.NonPublic).SetValue(null,CultureInfo.GetCultureInfo(name)); }
        public void Dispose() { Set(culture.Name); Thread.CurrentThread.CurrentCulture=hostCulture; Thread.CurrentThread.CurrentUICulture=hostUiCulture; }
    }
    internal sealed class EmptyResourceManager : ResourceManager
    {
        public override string GetString(string name,CultureInfo culture) { return null; }
    }
}
