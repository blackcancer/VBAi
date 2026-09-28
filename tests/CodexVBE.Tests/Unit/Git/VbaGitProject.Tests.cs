namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Reflection;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbaGitProjectTests
    {
        [TestMethod]
        public void OpenModuleGuardsAndLineSelectionPreserveLinkedProjectIdentity()
        {
            var host = new ProjectFixture(); var component = new ComponentFixture("Module1", 1, Code("Module1")); host.VBComponents.Items.Add(component);
            var adapter = Adapter(host); adapter.OpenModule("Module1", 0); Assert.IsTrue(component.CodeModule.CodePane.Shown); Assert.AreEqual(1, component.CodeModule.CodePane.Line);
            adapter.OpenModule("Module1", 7); Assert.AreEqual(7, component.CodeModule.CodePane.Line);
            Assert.ThrowsException<InvalidOperationException>(() => adapter.OpenModule("../bad"));
            host.FileName += ".changed"; Assert.ThrowsException<InvalidOperationException>(() => adapter.Capture()); host.FileName = host.FileName.Replace(".changed", "");
            host.Mode = 1; Assert.ThrowsException<InvalidOperationException>(() => adapter.Capture()); host.Mode = 2;
            host.Protection = 1; Assert.ThrowsException<InvalidOperationException>(() => adapter.Capture()); host.Protection = 0;
            host.References.Add(new ReferenceFixture()); Assert.IsTrue(adapter.Capture().Manifest.References.Contains("C000"));
            host.References[0].IsBroken = true; Assert.ThrowsException<InvalidOperationException>(() => adapter.Capture());
        }

        [TestMethod]
        public void ModuleClassFormResourcesAndDocumentReplacementMatrix()
        {
            var host = new ProjectFixture();
            host.VBComponents.Items.Add(new ComponentFixture("Module1", 1, Code("Module1")));
            host.VBComponents.Items.Add(new ComponentFixture("ThisWorkbook", 100, "Option Explicit\n"));
            host.VBComponents.Items.Add(new ComponentFixture("Form1", 3, Code("Form1")) { Resource = new byte[] { 0, 1 } });
            host.VBComponents.Items.Add(new ComponentFixture("EmptyForm", 3, Code("EmptyForm")));
            var adapter = Adapter(host); var initial = adapter.Capture(); adapter.Apply(initial, initial);
            Assert.IsTrue(initial.Manifest.Components[1].HasResources);
            var targetHost = new ProjectFixture();
            targetHost.VBComponents.Items.Add(new ComponentFixture("Module1", 2, Code("Module1", "' class replacement\n")));
            targetHost.VBComponents.Items.Add(new ComponentFixture("ThisWorkbook", 100, ""));
            targetHost.VBComponents.Items.Add(new ComponentFixture("Form1", 3, Code("Form1")));
            targetHost.VBComponents.Items.Add(new ComponentFixture("EmptyForm", 3, Code("EmptyForm")));
            var target = Adapter(targetHost).Capture(); int mutations = 0;
            adapter.Apply(target, initial, () => mutations++); Assert.AreEqual(1, mutations); Assert.IsTrue(adapter.Capture().SameAs(target));
            Assert.AreEqual("", host.VBComponents.Item("ThisWorkbook").CodeModule.Text);
            targetHost.VBComponents.Item("ThisWorkbook").CodeModule.Text = "Option Explicit\n' inserted into an empty document\n";
            var nonempty = Adapter(targetHost).Capture(); adapter.Apply(nonempty, target); Assert.IsTrue(adapter.Capture().SameAs(nonempty));
            targetHost.VBComponents.Item("Form1").Resource = new byte[] { 0, 2 };
            var resources = Adapter(targetHost).Capture(); adapter.Apply(resources, nonempty); Assert.IsTrue(adapter.Capture().SameAs(resources));
            targetHost.VBComponents.Item("Form1").Resource = new byte[] { 0, 3 };
            var changedResources = Adapter(targetHost).Capture(); adapter.Apply(changedResources, resources); Assert.IsTrue(adapter.Capture().SameAs(changedResources));
            host.VBComponents.Item("ThisWorkbook").CodeModule.Text = "' live mutation\n";
            Assert.ThrowsException<InvalidOperationException>(() => adapter.Apply(target, changedResources));
        }

        [TestMethod]
        public void ImportedNameTypeAndContentMismatchRemainVisibleWithoutAutomaticRetry()
        {
            foreach (int mode in new[] { 0, 1, 2 })
            {
                var host = new ProjectFixture(); host.VBComponents.Items.Add(new ComponentFixture("Module1", 1, Code("Module1")));
                var adapter = Adapter(host); var initial = adapter.Capture();
                var targetHost = new ProjectFixture(); targetHost.VBComponents.Items.Add(new ComponentFixture("Module1", 1, Code("Module1", "' replacement\n")));
                host.VBComponents.WrongName = mode == 0; host.VBComponents.WrongType = mode == 1; host.VBComponents.LoseText = mode == 2;
                Assert.ThrowsException<InvalidOperationException>(() => adapter.Apply(Adapter(targetHost).Capture(), initial));
                Assert.AreEqual(1, host.VBComponents.Items.Count);
            }
        }

        [TestMethod]
        public void ReferencesAndHostDocumentModulesMustMatchBeforeImport()
        {
            var host = new ProjectFixture(); host.VBComponents.Items.Add(new ComponentFixture("ThisWorkbook", 100, "Option Explicit\n"));
            var adapter = Adapter(host); var initial = adapter.Capture();
            var target = new ProjectFixture(); target.VBComponents.Items.Add(new ComponentFixture("ThisWorkbook", 100, "Option Explicit\n"));
            target.References.Add(new ReferenceFixture());
            Assert.ThrowsException<InvalidOperationException>(() => adapter.Apply(Adapter(target).Capture(), initial));
            target.References.Clear(); target.VBComponents.Item("ThisWorkbook").Name = "AnotherDocument";
            Assert.ThrowsException<InvalidOperationException>(() => adapter.Apply(Adapter(target).Capture(), initial));
            Assert.IsTrue(adapter.Capture().SameAs(initial));
        }

        [TestMethod]
        public void ScratchCleanupToleratesLockedAndReadonlyOwnTemporaryFiles()
        {
            var type = typeof(VbaGitProject).GetNestedType("Scratch", BindingFlags.NonPublic);
            foreach (bool locked in new[] { true, false })
            {
                var scratch = (IDisposable)Activator.CreateInstance(type, true);
                string path = (string)type.GetField("Path", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(scratch);
                string parent = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "GitTemporary"));
                Assert.AreEqual(parent, Directory.GetParent(Path.GetFullPath(path)).FullName);
                Assert.IsTrue(Guid.TryParseExact(Path.GetFileName(path), "N", out _));
                string file = Path.Combine(path, "owned-cleanup-fixture.txt"); File.WriteAllText(file, "Disposable coverage fixture");
                try
                {
                    if (locked) using (File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None)) scratch.Dispose();
                    else { File.SetAttributes(file, FileAttributes.ReadOnly); scratch.Dispose(); }
                    Assert.IsTrue(File.Exists(file));
                }
                finally { File.SetAttributes(file, FileAttributes.Normal); File.Delete(file); Directory.Delete(path); }
            }
        }
    }
}
