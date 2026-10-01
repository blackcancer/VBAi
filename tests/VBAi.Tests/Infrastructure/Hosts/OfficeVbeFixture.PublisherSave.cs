using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Saves only the reviewed synthetic Publisher project through the VBE Save Host Document control.</summary>
        internal void SaveReviewedPublisherProject(string moduleName, IDictionary<string, string> expectedSources)
        {
            RequireUsableOwnedHost();
            Assert.AreEqual("Publisher", Kind);
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
            Assert.IsFalse(NativeExecutionUnsettled, "Unsettled native execution forbids the final Publisher save.");
            Assert.IsTrue(expectedSources != null && expectedSources.ContainsKey(moduleName), "The synthetic module must be reviewed.");
            var reviewed = new Dictionary<string, string>(expectedSources, StringComparer.OrdinalIgnoreCase);
            var acquisitions = new List<object>();
            Func<object, object> acquire = item => { if (item != null) acquisitions.Add(item); return item; };
            bool saveEntered = false;
            try
            {
                object project = acquire(((dynamic)document).VBProject);
                object editor = acquire(((dynamic)application).VBE);
                object components = acquire(((dynamic)project).VBComponents);
                object component = acquire(((dynamic)components).Item(moduleName));
                object code = acquire(((dynamic)component).CodeModule);
                object pane = acquire(((dynamic)code).CodePane);
                object window = acquire(((dynamic)pane).Window);
                object bars = acquire(((dynamic)editor).CommandBars);

                Action<bool> validate = requirePane => {
                    RequireUsableOwnedHost();
                    Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited, "The original Publisher process must remain alive.");
                    Assert.IsTrue(File.Exists(DocumentPath) && string.Equals(Path.GetExtension(DocumentPath), ".pub", StringComparison.OrdinalIgnoreCase));
                    Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode));
                    Assert.AreEqual(0, Convert.ToInt32(((dynamic)project).Protection));
                    Assert.IsFalse((bool)((dynamic)document).ReadOnly);
                    Assert.AreEqual(Path.GetFullPath(DocumentPath), Path.GetFullPath((string)((dynamic)document).FullName), true);
                    object documents = acquire(((dynamic)application).Documents);
                    Assert.AreEqual(1, Convert.ToInt32(((dynamic)documents).Count), "Only the owned synthetic publication may be saved.");
                    Assert.IsTrue(SamePublisherSaveIdentity(document, acquire(((dynamic)documents)[1])));
                    Assert.IsTrue(SamePublisherSaveIdentity(document, acquire(((dynamic)application).ActiveDocument)));
                    Assert.IsTrue(SamePublisherSaveIdentity(project, acquire(((dynamic)document).VBProject)));
                    Assert.IsTrue(SamePublisherSaveIdentity(editor, acquire(((dynamic)application).VBE)));
                    Assert.IsTrue(SamePublisherSaveIdentity(editor, acquire(((dynamic)project).VBE)));
                    Assert.IsTrue(SamePublisherSaveIdentity(components, acquire(((dynamic)project).VBComponents)));
                    object hostWindow = acquire(((dynamic)application).ActiveWindow);
                    Assert.AreEqual((uint)ProcessId, ReadPublisherWindowOwner(new IntPtr(Convert.ToInt64(((dynamic)hostWindow).hWnd))));
                    object mainWindow = acquire(((dynamic)editor).MainWindow);
                    Assert.AreEqual((uint)ProcessId, ReadPublisherWindowOwner(new IntPtr(Convert.ToInt64(((dynamic)mainWindow).HWnd))));
                    var actual = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    int count = Convert.ToInt32(((dynamic)components).Count);
                    for (int index = 1; index <= count; index++)
                    {
                        object item = acquire(((dynamic)components).Item(index));
                        object source = acquire(((dynamic)item).CodeModule);
                        int lines = Convert.ToInt32(((dynamic)source).CountOfLines);
                        actual.Add((string)((dynamic)item).Name, lines == 0 ? "" : (string)((dynamic)source).Lines[1, lines]);
                    }
                    RequirePublisherSaveSources(reviewed, actual);
                    if (requirePane)
                    {
                        Assert.IsTrue(SamePublisherSaveIdentity(project, acquire(((dynamic)editor).ActiveVBProject)));
                        object activePane = acquire(((dynamic)editor).ActiveCodePane);
                        Assert.IsTrue(SamePublisherSaveIdentity(pane, activePane));
                        object activeCode = acquire(((dynamic)activePane).CodeModule);
                        Assert.IsTrue(SamePublisherSaveIdentity(component, acquire(((dynamic)activeCode).Parent)));
                        Assert.IsTrue(SamePublisherSaveIdentity(window, acquire(((dynamic)editor).ActiveWindow)));
                        Assert.AreEqual(0, Convert.ToInt32(((dynamic)window).Type));
                        Assert.IsTrue((bool)((dynamic)window).Visible);
                        int firstLine, firstColumn, lastLine, lastColumn;
                        ((dynamic)activePane).GetSelection(out firstLine, out firstColumn, out lastLine, out lastColumn);
                        Assert.IsTrue(firstLine == 1 && lastLine == 1 && firstColumn == 1 && lastColumn == 1,
                            "The reviewed synthetic module selection changed.");
                    }
                };

                validate(false);
                ((dynamic)editor).ActiveVBProject = project;
                ((dynamic)pane).Show();
                ((dynamic)pane).SetSelection(1, 1, 1, 1);
                ((dynamic)window).SetFocus();
                validate(true);
                object control = acquire(((dynamic)bars).FindControl(1, 3));
                string caption = RequirePublisherSaveControl(control);
                validate(true);
                object currentControl = acquire(((dynamic)bars).FindControl(1, 3));
                Assert.IsTrue(SamePublisherSaveIdentity(control, currentControl), "The native VBE save control changed.");
                Assert.AreEqual(caption, RequirePublisherSaveControl(currentControl), "The native VBE save command changed.");
                validate(true);

                // Mark the boundary before emission. A throw or an incomplete readback never permits retry or Quit.
                NativeExecutionUnsettled = true;
                saveEntered = true;
                steps.Add(new { PublisherFinalSave = "EnteringNativeControl", ControlId = 3, ProcessId, DocumentPath, Module = moduleName });
                FlushAdapterEvidence();
                ((dynamic)control).Execute();
                validate(true);
                Assert.IsTrue((bool)((dynamic)project).Saved, "VBE Save returned without a saved VBA project.");
                Assert.IsTrue((bool)((dynamic)document).Saved, "VBE Save returned without a saved publication.");
                var persistence = Data("project_persistence_status", "Project", DocumentPath);
                Assert.AreEqual("Publisher", persistence["Host"]);
                Assert.AreEqual(ProcessId, Convert.ToInt32(persistence["OwnerProcessId"]));
                Assert.AreEqual(true, persistence["IdentityVerified"]);
                Assert.AreEqual(Path.GetFullPath(DocumentPath), Path.GetFullPath((string)persistence["HostPath"]), true);
                Assert.AreEqual(true, persistence["HostSaved"]);
                Assert.AreEqual(true, persistence["ProjectSaved"]);
                validate(true);
                Assert.IsTrue((bool)((dynamic)project).Saved && (bool)((dynamic)document).Saved,
                    "The exact project's saved state changed during bridge readback.");
                steps.Add(new { PublisherFinalSave = "Verified", ControlId = 3, ProcessId, DocumentPath, Module = moduleName,
                    ProjectSaved = true, HostSaved = true, Persistence = persistence });
                FlushAdapterEvidence();
                NativeExecutionUnsettled = false;
            }
            catch (Exception error)
            {
                if (saveEntered)
                {
                    NativeExecutionUnsettled = true;
                    RetainUncertainOffice();
                    steps.Add(new { PublisherFinalSave = "UnverifiedAfterEmission", Error = error.ToString(), ProcessId, DocumentPath,
                        SaveRetried = false, QuitPermitted = false });
                }
                else steps.Add(new { PublisherFinalSave = "RefusedBeforeEmission", Error = error.ToString(), ProcessId, DocumentPath });
                throw;
            }
            finally
            {
                if (NativeExecutionUnsettled || commandContainment.Pending || commandContainment.Uncertain || hostTeardownRefused)
                    retainedDiagnosticReferences.AddRange(acquisitions);
                else
                    for (int index = acquisitions.Count - 1; index >= 0; index--)
                        if (Marshal.IsComObject(acquisitions[index])) Marshal.ReleaseComObject(acquisitions[index]);
            }
        }

        internal static void RequirePublisherSaveSources(IDictionary<string, string> expected, IDictionary<string, string> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count, "The publication's module inventory changed.");
            foreach (var source in expected)
            {
                Assert.IsTrue(actual.TryGetValue(source.Key, out string text), "The publication's module identity changed: " + source.Key);
                Assert.AreEqual(CanonicalSource(source.Value), CanonicalSource(text), "Unreviewed Publisher source: " + source.Key);
            }
        }

        private static string RequirePublisherSaveControl(object control)
        {
            Assert.IsNotNull(control, "The native VBE Save Host Document command is unavailable.");
            dynamic item = control;
            Assert.AreEqual(3, Convert.ToInt32(item.Id));
            Assert.AreEqual(1, Convert.ToInt32(item.Type));
            Assert.IsTrue((bool)item.BuiltIn && (bool)item.Enabled, "The built-in VBE save command must be enabled.");
            Assert.IsTrue(string.IsNullOrEmpty((string)item.OnAction), "A custom save action is refused.");
            return (string)item.Caption;
        }

        private static bool SamePublisherSaveIdentity(object first, object second)
        {
            if (first == null || second == null) return false;
            if (!Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return ReferenceEquals(first, second);
            IntPtr left = IntPtr.Zero, right = IntPtr.Zero;
            try { left = Marshal.GetIUnknownForObject(first); right = Marshal.GetIUnknownForObject(second); return left == right; }
            finally { if (right != IntPtr.Zero) Marshal.Release(right); if (left != IntPtr.Zero) Marshal.Release(left); }
        }
    }
}
