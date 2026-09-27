using System;
using System.Drawing;
using System.Reflection;
using CodexVBE;
namespace CodexVBE.Tests.Infrastructure
{
    internal sealed class ThemeScope : IDisposable
    {
        private readonly ThemeChoice choice = UiTheme.Choice;
        private readonly string file = UiTheme.FileName;
        private readonly Func<bool> contrast = UiTheme.HighContrast;
        private readonly Func<Color> window = UiTheme.WindowColor;
        private readonly Func<object> system = UiTheme.ReadSystemTheme;
        private readonly Func<string,string> read = UiTheme.ReadTheme;
        private readonly Action<string,string> write = UiTheme.WriteTheme;
        internal ThemeScope(bool dark = false)
        {
            UiTheme.HighContrast = () => false;
            UiTheme.WriteTheme = (p,v) => { };
            SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
        }
        internal static void SetChoice(ThemeChoice value) { typeof(UiTheme).GetProperty("Choice", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null,value); }
        public void Dispose() { SetChoice(choice); UiTheme.FileName=file; UiTheme.HighContrast=contrast; UiTheme.WindowColor=window; UiTheme.ReadSystemTheme=system; UiTheme.ReadTheme=read; UiTheme.WriteTheme=write; }
    }
    internal sealed class DesignContext : System.ComponentModel.LicenseContext
    {
        public override System.ComponentModel.LicenseUsageMode UsageMode { get { return System.ComponentModel.LicenseUsageMode.Designtime; } }
    }
    internal static class UiInvoke
    {
        internal static object Call(Type type, string method, object target, params object[] args)
        {
            return type.GetMethod(method, BindingFlags.Static | BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target,args);
        }
    }
}
