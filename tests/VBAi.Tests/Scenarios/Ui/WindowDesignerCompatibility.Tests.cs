using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.ComponentModel.Design.Serialization;
using System.Linq;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Scenarios.Ui
{
    [TestClass, TestCategory("Unit")]
    public sealed class WindowDesignerCompatibilityTests
    {
        [STATestMethod]
        public void UnsitedDesignerPreviewHonorsTheActiveDesignTimeLicenseContext()
        {
            var previous = LicenseManager.CurrentContext;
            try
            {
                LicenseManager.CurrentContext = new DesigntimeLicenseContext();
                foreach (var type in new[] { typeof(ModernEditorWindow), typeof(UpdateWindow), typeof(UpdateProgressWindow) })
                using (var form = (Form)Activator.CreateInstance(type))
                {
                    Assert.IsNull(form.Site, "This case exercises the license guard rather than a Designer site.");
                    if (form is UpdateWindow updates) updates.ReadPreferences = () => { Assert.Fail("Preview read update preferences"); return null; };
                    type.GetMethod("OnShown", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        .Invoke(form, new object[] { EventArgs.Empty });
                    if (form is ModernEditorWindow editor) Assert.IsNull(editor.Browser);
                }
            }
            finally { LicenseManager.CurrentContext = previous; }
        }

        [STATestMethod]
        public void EditorUpdatesGitAndSettingsHaveEditableSerializableDesignerControls()
        {
            foreach (var type in new[] { typeof(ModernEditorWindow), typeof(UpdateWindow), typeof(UpdateProgressWindow), typeof(GitWindow), typeof(LlmSettingsWindow) }) {
                var context = new DesigntimeLicenseContext();
                using (var subject = (Form)LicenseManager.CreateWithContext(type, context))
                using (var surface = new DesignSurface(typeof(Form))) {
                    var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                    var root = (Form)host.RootComponent;
                    foreach (Control child in subject.Controls.Cast<Control>().ToArray()) root.Controls.Add(child);
                    Assert.AreEqual(0, surface.LoadErrors.Count, type.Name);
                    var components = type.GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Select(f => new { f.Name, Component = f.GetValue(subject) as IComponent })
                        .Where(x => x.Component != null && (x.Component is Control || x.Component is Timer || x.Component is ToolTip)).ToArray();
                    Assert.IsTrue(components.Length > 0, type.Name);
                    foreach (var entry in components) {
                        if (entry.Component.Site != null) entry.Component.Site.Container.Remove(entry.Component);
                        host.Container.Add(entry.Component, entry.Name);
                        Assert.IsNotNull(host.GetDesigner(entry.Component), type.Name + "/" + entry.Name);
                        if (entry.Component is Control control) {
                            var prop = TypeDescriptor.GetProperties(control)["AccessibleDescription"];
                            prop.SetValue(control, "Editable Designer component");
                            var service = new CodeDomComponentSerializationService(host);
                            using (var store = service.CreateStore()) {
                                service.SerializeAbsolute(store, control); store.Close();
                                Assert.AreEqual(0, store.Errors.Count, type.Name + "/" + entry.Name);
                                using (var copies = new Container()) {
                                    var copy = service.Deserialize(store, copies).OfType<Control>().First(x => x.Name == control.Name);
                                    Assert.AreEqual("Editable Designer component", copy.AccessibleDescription);
                                }
                            }
                        }
                    }
                    if (subject is ModernEditorWindow editor) Assert.IsNull(editor.Browser, "Designer must not start Chromium");
                }
            }
        }
        [STATestMethod]
        public void DesignerPreviewDoesNotInitializeMonacoOrReadUpdatePreferences()
        {
            foreach (var type in new[] { typeof(ModernEditorWindow), typeof(UpdateWindow), typeof(UpdateProgressWindow) })
                using (var surface = new DesignSurface(type)) {
                    var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                    var form = (Form)host.RootComponent;
                    if (form is UpdateWindow updates) updates.ReadPreferences = () => { Assert.Fail("Designer read update preferences"); return null; };
                    type.GetMethod("OnShown", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        .Invoke(form, new object[] { EventArgs.Empty });
                    if (form is ModernEditorWindow editor) Assert.IsNull(editor.Browser);
                }
        }
    }
}
