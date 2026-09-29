using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Reflection;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Infrastructure
{
    internal static class TranscriptFixture
    {
        internal static void LayoutAndLifecycle<T>() where T : UserControl, new()
        {
            using (var theme = new ThemeScope())
            using (var surface = new DesignSurface(typeof(T)))
            {
                Assert.IsTrue(surface.IsLoaded); Assert.AreEqual(0, surface.LoadErrors.Count);
                var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                var root = (T)host.RootComponent; Assert.IsTrue(root.Controls.Count > 0);
                root.Width = 420; root.PerformLayout(); Assert.IsNotNull(host.GetDesigner(root));
            }
            using (var theme = new ThemeScope())
            {
                var view = new T(); Assert.AreEqual(typeof(T).Name, view.Name);
                var containerField = typeof(T).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
                var container = (IContainer)containerField.GetValue(view);
                Assert.IsNotNull(container);
                var component = new Component(); bool disposed = false; component.Disposed += (s,e) => disposed = true; container.Add(component);
                Dispose(view, false); Assert.IsFalse(disposed);
                view.Dispose(); Assert.IsTrue(disposed); Assert.IsTrue(view.IsDisposed);
                var absent = new T(); var owned = (IContainer)containerField.GetValue(absent); containerField.SetValue(absent, null);
                absent.Dispose(); Assert.IsTrue(absent.IsDisposed); owned.Dispose();
            }
        }
        internal static void Dispose(Control view, bool disposing) { view.GetType().GetMethod("Dispose", BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(bool) }, null).Invoke(view, new object[] { disposing }); }
        internal static void Set(object target, string field, object value) { target.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value); }
        internal static void Event(Control control, string method, EventArgs args) { control.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { args }); }
    }
}