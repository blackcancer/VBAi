using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks Office document adapters with injectable probes; native qualification is independent.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed partial class VbeOtherHostPersistenceTests
    {
        [DataTestMethod]
        [DataRow("HostPath"), DataRow("ProjectPath"), DataRow("HostSaved"), DataRow("ProjectSaved")]
        [DataRow("FileExists"), DataRow("FileLength"), DataRow("FileFormat"), DataRow("SourceSha256"), DataRow("ProjectIdentity")]
        public void SaveVerificationReportsExactFailedGuardWithoutRepeatingMutation(string guard)
        {
            var f = new Fixture { Kind = "PowerPoint" };
            f.Observation.Path = @"C:\fixture\Document.pptm";
            f.Project.FileName = f.Observation.Path; f.Observation.Format = 25;
            f.AfterInvocation = () =>
            {
                f.Observation.Format = 25;
                switch (guard)
                {
                    case "HostPath": f.Observation.Path = @"C:\fixture\Other.pptm"; f.Project.Saved = false; break;
                    case "ProjectPath": f.Project.FileName = @"C:\fixture\Other.pptm"; break;
                    case "HostSaved": f.Observation.Saved = false; break;
                    case "ProjectSaved": f.Project.Saved = false; break;
                    case "FileExists": f.Exists = false; break;
                    case "FileLength": f.Bytes = 0; break;
                    case "FileFormat": f.Observation.Format = 99; break;
                    case "SourceSha256": f.Component.CodeModule.Source += "\r\n' Changed during save"; break;
                    case "ProjectIdentity": f.Identity = false; break;
                }
            };
            dynamic result = f.Service.SaveOtherHost(f.Request(), false, f);
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
            Assert.IsTrue((bool)result.MutationInvoked); Assert.AreEqual(1, f.Attempts);
            Assert.AreEqual("Save verification failed: " + guard + ".", (string)result.Reason);
        }

        [DataTestMethod, DataRow("Access", ".accdb", 12), DataRow("Publisher", ".pub", 1)]
        public void ExistingAccessPublisherSaveKeepsNullableHostStateAndNeverRetries(string kind, string extension, int format)
        {
            var f = new Fixture { Kind = kind };
            f.Observation.Path = @"C:\fixture\Document" + extension;
            f.Project.FileName = f.Observation.Path; f.Observation.Format = format;
            f.Observation.Saved = kind == "Access" ? (bool?)null : true;
            f.AfterInvocation = () => { f.Observation.Format = format; f.Observation.Saved = kind == "Access" ? (bool?)null : true; };
            dynamic result = f.Service.SaveOtherHost(f.Request(), false, f);
            Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain); Assert.AreEqual(1, f.Attempts);
            Assert.AreEqual(kind == "Access" ? (bool?)null : true, (bool?)result.HostSaved);
            Assert.IsTrue((bool)result.NativeFileFormatVerified); Assert.IsFalse((bool)result.PersistenceReopenVerified);
            f.AfterInvocation = () => { f.Observation.Format = format; f.Project.Saved = false; };
            result = f.Service.SaveOtherHost(f.Request(), false, f);
            Assert.IsTrue((bool)result.Uncertain); Assert.AreEqual(2, f.Attempts);
        }

        [DataTestMethod, DataRow("Access", ".accdb", 12), DataRow("Publisher", ".pub", 1)]
        public void AccessPublisherRefuseStalePathsFormatsReadonlyAndFirstSaveBeforeMutation(string kind, string extension, int format)
        {
            foreach (string fault in new[] { "path", "project", "format", "readonly", "pid", "duplicate", "saveas" })
            {
                var f = new Fixture { Kind = kind };
                f.Observation.Path = @"C:\fixture\Document" + extension; f.Project.FileName = f.Observation.Path; f.Observation.Format = format;
                var request = f.Request();
                if (fault == "path") request.ExpectedHostPath = @"C:\fixture\Other" + extension;
                if (fault == "project") f.Project.FileName = @"C:\fixture\Other" + extension;
                if (fault == "format") f.Observation.Format = 99;
                if (fault == "readonly") f.Observation.ReadOnly = true;
                if (fault == "pid") f.Owner = 43;
                if (fault == "duplicate") f.Items.Add(f);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveOtherHost(request, fault == "saveas", f), fault);
                Assert.AreEqual(0, f.Attempts, fault);
            }
        }

        [DataTestMethod, DataRow("Access"), DataRow("Publisher")]
        public void NativeAccessPublisherAssociateInjectedVbeAndExactPathWithoutDocumentVBProject(string kind)
        {
            var editor = new PathEditor(); var project = new PathProject { VBE = editor, FileName = kind == "Access" ? @"C:\fixture\Owned.accdb" : @"C:\fixture\Owned.pub" };
            editor.VBProjects.Add(project); editor.ActiveVBProject = project;
            var app = new PathApplication(); var doc = new PathDocument { Application = app, FullName = project.FileName };
            app.CurrentProject = doc; app.Documents.Add(doc);
            var native = new VbeProjectComponents.NativeOtherHostProbe
            {
                ReadHostKind = () => kind,
                ReadOwner = h => h.ToInt64() == 77 ? (uint)System.Diagnostics.Process.GetCurrentProcess().Id : 0,
                ReadIdentity = ReferenceEquals,
                ReadActiveApplication = name => app
            };
            native.BindProject(project);
            Assert.AreSame(app, native.Application()); Assert.AreSame(doc, native.Documents(app)[0]);
            Assert.AreSame(project, native.DocumentProject(doc));
            native.BindDocument(doc);
            Assert.AreEqual(kind == "Access" ? (bool?)null : true, native.State(doc).Saved);
            if (kind == "Access")
            {
                native.PrepareSave(doc); native.Save(doc, false, project.FileName, 12); Assert.AreEqual(1, editor.CommandBars.Control.Executions);
                editor.ActiveVBProject = new object();
                Assert.ThrowsException<InvalidOperationException>(() => native.PrepareSave(doc));
                Assert.AreEqual(1, editor.CommandBars.Control.Executions);
                editor.ActiveVBProject = project; editor.CommandBars.Control.Enabled = false;
                Assert.ThrowsException<InvalidOperationException>(() => native.PrepareSave(doc));
            }
            else { native.Save(doc, false, project.FileName, 1); Assert.AreEqual(1, doc.SaveCalls); }
            editor.VBProjects.Add(new PathProject { VBE = editor, FileName = project.FileName });
            Assert.ThrowsException<InvalidOperationException>(() => native.DocumentProject(doc));
            editor.VBProjects.RemoveAt(1); project.FileName = @"C:\fixture\Other.pub";
            Assert.ThrowsException<InvalidOperationException>(() => native.DocumentProject(doc));
        }

        public sealed class PathProject
        {
            private string fileName;
            public object VBE { get; set; }
            public Exception PathReadError { get; set; }
            public string FileName { get { if (PathReadError != null) throw PathReadError; return fileName; } set { fileName = value; } }
        }
        public sealed class PathWindow { public int HWnd => 77; public int hWnd => 77; }
        public sealed class PathEditor
        {
            public PathWindow MainWindow { get; } = new PathWindow();
            public System.Collections.Generic.List<object> VBProjects { get; } = new System.Collections.Generic.List<object>();
            public object ActiveVBProject { get; set; }
            public PathCommands CommandBars { get; } = new PathCommands();
        }
        public sealed class PathCommands
        {
            public PathControl Control { get; } = new PathControl();
            public object FindControl(int type, int id) { Assert.AreEqual(1, type); Assert.AreEqual(3, id); return Control; }
        }
        public sealed class PathControl
        {
            public int Id => 3; public int Type => 1; public bool BuiltIn => true; public bool Enabled { get; set; } = true;
            public int Executions { get; private set; }
            public void Execute() { Executions++; }
        }
        public sealed class PathDatabase { public bool Updatable => true; }
        public sealed class PathApplication
        {
            public PathWindow ActiveWindow { get; } = new PathWindow(); public int hWndAccessApp() => 77;
            public object CurrentProject { get; set; }
            public System.Collections.Generic.List<object> Documents { get; } = new System.Collections.Generic.List<object>();
            public object CurrentDb() => new PathDatabase();
        }
        public sealed class PathDocument
        {
            public object Application { get; set; }
            public string FullName { get; set; }
            public string Path => @"C:\fixture";
            public int FileFormat => 12; public int SaveFormat { get; set; } = 1; public bool Saved => true; public bool ReadOnly { get; set; }
            public int SaveCalls { get; private set; }
            public void Save() { SaveCalls++; }
        }

        [TestMethod]
        public void AccessVersionedRotFallbackStillRequiresTheCurrentProcessOwner()
        {
            var app = new PathApplication(); var attempts = new System.Collections.Generic.List<string>();
            uint owner = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            var native = new VbeProjectComponents.NativeOtherHostProbe
            {
                ReadHostKind = () => "Access",
                ReadHostMajorVersion = () => 16,
                ReadOwner = h => owner,
                ReadActiveApplication = name =>
                {
                    attempts.Add(name);
                    if (name == "Access.Application") throw new System.Runtime.InteropServices.COMException("Not registered", unchecked((int)0x800401E3));
                    return app;
                }
            };
            Assert.AreSame(app, native.Application());
            CollectionAssert.AreEqual(new[] { "Access.Application", "Access.Application.16" }, attempts);
            owner = 0;
            Assert.ThrowsException<InvalidOperationException>(() => native.Application());
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void PublisherWithoutProjectFileNameRequiresOneExactDocumentAndProjectBeforeSave(bool unavailablePath)
        {
            foreach (string fault in new[] { "none", "documents", "foreign document", "projects", "foreign project", "foreign vbe", "pid", "path", "format", "readonly" })
            {
                var editor = new PathEditor();
                var project = new PathProject
                {
                    VBE = editor,
                    FileName = "",
                    PathReadError = unavailablePath ? new DirectoryNotFoundException("Path not found") : null
                };
                editor.VBProjects.Add(project);
                var app = new PathApplication(); var doc = new PathDocument { Application = app, FullName = @"C:\fixture\Owned.pub" };
                app.Documents.Add(doc);
                uint owner = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                var native = new VbeProjectComponents.NativeOtherHostProbe
                {
                    ReadHostKind = () => "Publisher",
                    ReadOwner = h => owner,
                    ReadIdentity = ReferenceEquals,
                    ReadActiveApplication = name => app
                };
                native.BindProject(project); native.BindDocument(doc);
                Assert.AreSame(project, native.DocumentProject(doc));
                if (fault == "documents") app.Documents.Add(new PathDocument { Application = app, FullName = @"C:\fixture\Other.pub" });
                if (fault == "foreign document") app.Documents[0] = new PathDocument { Application = app, FullName = doc.FullName };
                if (fault == "projects") editor.VBProjects.Add(new PathProject { VBE = editor });
                if (fault == "foreign project") editor.VBProjects[0] = new PathProject { VBE = editor };
                if (fault == "foreign vbe") project.VBE = new PathEditor();
                if (fault == "pid") owner = 0;
                if (fault == "path") doc.FullName = @"C:\fixture\Changed.pub";
                if (fault == "format") doc.SaveFormat = 3;
                if (fault == "readonly") doc.ReadOnly = true;
                if (fault == "none")
                {
                    native.Save(doc, false, @"C:\fixture\Owned.pub", 1);
                    Assert.AreEqual(1, doc.SaveCalls); Assert.IsTrue(native.SaveInvocationStarted);
                }
                else
                {
                    Assert.ThrowsException<InvalidOperationException>(() => native.Save(doc, false, @"C:\fixture\Owned.pub", 1), fault);
                    Assert.AreEqual(0, doc.SaveCalls, fault); Assert.IsFalse(native.SaveInvocationStarted, fault);
                }
            }
        }

        [TestMethod]
        public void PublisherProbeRevalidatesEveryNativePreconditionBeforeInvokingSave()
        {
            foreach (string fault in new[] { "document", "collection", "pid", "project", "path", "format", "readonly" })
            {
                var editor = new PathEditor(); var project = new PathProject { VBE = editor, FileName = @"C:\fixture\Owned.pub" };
                editor.VBProjects.Add(project);
                var app = new PathApplication(); var doc = new PathDocument { Application = app, FullName = project.FileName };
                app.Documents.Add(doc);
                uint owner = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                var native = new VbeProjectComponents.NativeOtherHostProbe
                {
                    ReadHostKind = () => "Publisher",
                    ReadOwner = h => owner,
                    ReadIdentity = ReferenceEquals,
                    ReadActiveApplication = name => app
                };
                native.BindProject(project); native.BindDocument(doc);
                Assert.AreSame(project, native.DocumentProject(doc));
                string destination = project.FileName;
                if (fault == "document") native.BindDocument(new PathDocument { Application = app, FullName = destination });
                if (fault == "collection") app.Documents.Clear();
                if (fault == "pid") owner = 0;
                if (fault == "project") editor.VBProjects.Clear();
                if (fault == "path") { project.FileName = @"C:\fixture\Changed.pub"; doc.FullName = project.FileName; }
                if (fault == "format") doc.SaveFormat = 3;
                if (fault == "readonly") doc.ReadOnly = true;
                Assert.ThrowsException<InvalidOperationException>(() => native.Save(doc, false, destination, 1), fault);
                Assert.AreEqual(0, doc.SaveCalls, fault); Assert.IsFalse(native.SaveInvocationStarted, fault);
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void WordTemporaryVbaStorageDoesNotReplaceMatchedDocumentPath(bool temporaryPathBeforeSave)
        {
            var fixture = new Fixture();
            if (temporaryPathBeforeSave) fixture.Project.FileName = @"C:\fixture\~WRL0002.tmp";
            else fixture.Project.PathReadError = new DirectoryNotFoundException("Path not found");
            fixture.AfterInvocation = () =>
            {
                fixture.Project.PathReadError = null;
                fixture.Project.FileName = @"C:\fixture\~WRL0002.tmp";
            };
            dynamic result = fixture.Service.SaveOtherHost(fixture.Request(), false, fixture);
            Assert.IsTrue((bool)result.Verified);
            Assert.IsFalse((bool)result.Uncertain);
            Assert.AreEqual(@"C:\fixture\Document.docm", (string)result.HostPath);
            Assert.AreEqual(1, fixture.Attempts);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void PowerPointStillRequiresItsProjectPathBeforeAndAfterSaving(bool temporaryPathBeforeSave)
        {
            var fixture = new Fixture { Kind = "PowerPoint" };
            fixture.Observation.Path = @"C:\fixture\Document.pptm"; fixture.Observation.Format = null;
            fixture.Project.FileName = temporaryPathBeforeSave ? @"C:\fixture\~WRL0002.tmp" : fixture.Observation.Path;
            var request = fixture.Request();
            fixture.AfterInvocation = () => fixture.Project.FileName = @"C:\fixture\~WRL0002.tmp";
            if (temporaryPathBeforeSave)
            {
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveOtherHost(request, false, fixture));
                Assert.AreEqual(0, fixture.Attempts);
            }
            else
            {
                dynamic result = fixture.Service.SaveOtherHost(request, false, fixture);
                Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
                Assert.AreEqual(1, fixture.Attempts);
            }
        }

        [TestMethod]
        public void WordSaveUsesMatchedDocumentIdentityWhenItsProjectFilenameIsUnavailable()
        {
            foreach (Exception error in new Exception[] { new DirectoryNotFoundException("Path not found"),
                new System.Runtime.InteropServices.COMException("Path not found", unchecked((int)0x800A004C)) })
            {
                var fixture = new Fixture(); fixture.Project.PathReadError = error;
                dynamic result = fixture.Service.SaveOtherHost(fixture.Request(), false, fixture);
                Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
                Assert.AreEqual(1, fixture.Attempts); Assert.AreEqual(fixture.Observation.Path, (string)result.HostPath);
            }
        }

        [TestMethod]
        public void UnavailableWordProjectFilenameNeverBypassesNativeIdentityPathOrFormatGuards()
        {
            foreach (string fault in new[] { "path", "identity", "pid", "format", "duplicate" })
            {
                var fixture = new Fixture(); fixture.Project.PathReadError = new DirectoryNotFoundException("Path not found");
                var request = fixture.Request();
                if (fault == "path") request.ExpectedHostPath = @"C:\fixture\Other.docm";
                if (fault == "identity") fixture.Identity = false;
                if (fault == "pid") fixture.Owner = 43;
                if (fault == "format") fixture.Observation.Format = 12;
                if (fault == "duplicate") fixture.Items.Add(fixture);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveOtherHost(request, false, fixture), fault);
                Assert.AreEqual(0, fixture.Attempts, fault);
            }
            var denied = new Fixture { Kind = "PowerPoint" }; denied.Observation.Path = @"C:\fixture\Document.pptm"; denied.Observation.Format = null;
            denied.Project.PathReadError = new System.Runtime.InteropServices.COMException("Access denied", unchecked((int)0x80070005));
            Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => denied.Service.SaveOtherHost(denied.Request(), false, denied));
            Assert.AreEqual(0, denied.Attempts);
            var powerpoint = new Fixture { Kind = "PowerPoint" }; powerpoint.Observation.Path = @"C:\fixture\Document.pptm"; powerpoint.Observation.Format = null;
            powerpoint.Project.PathReadError = new DirectoryNotFoundException("Path not found");
            Assert.ThrowsException<InvalidOperationException>(() => powerpoint.Service.SaveOtherHost(powerpoint.Request(), false, powerpoint));
            Assert.AreEqual(0, powerpoint.Attempts);
        }

        [TestMethod]
        public void RemainingPersistenceGuardsRejectEveryPreflightTransitionWithoutInvokingSave()
        {
            var initial = new Fixture();
            foreach (Request invalid in new Request[] { null, new Request(), new Request { ExpectedProjectVersion = " " } })
                Assert.ThrowsException<ArgumentException>(() => initial.Service.SaveOtherHost(invalid, false, initial));
            Assert.IsFalse(initial.Service.SupportsOtherHost);
            dynamic unavailable = initial.Service.OtherHostPersistence("P"); Assert.IsFalse((bool)unavailable.HostAvailable);
            Assert.ThrowsException<InvalidOperationException>(() => initial.Service.SaveOtherHost(initial.Request(), false));
            initial.Project.Protection = 1;
            Assert.ThrowsException<InvalidOperationException>(() => initial.Service.SaveOtherHost(initial.Request(), false, initial));
            foreach (string fault in new[] { "null probe", "zero pid", "null documents", "oversized documents", "project path", "identity", "readonly", "path", "format", "source" })
            {
                var f = new Fixture(); var request = f.Request(fault == "project path");
                if (fault == "zero pid") f.ProcessId = 0;
                if (fault == "null documents") f.NullDocuments = true;
                if (fault == "oversized documents") for (int i = 0; i < 1000; i++) f.Items.Add(f);
                if (fault == "project path") f.Project.FileName = @"C:\fixture\Previous.docm";
                if (fault == "identity") f.ChangeIdentity = true;
                f.BeforeSecondState = () =>
                {
                    if (fault == "readonly") f.Observation.ReadOnly = true;
                    if (fault == "path") f.Observation.Path = @"C:\fixture\Changed.docm";
                    if (fault == "format") f.Observation.Format = 15;
                    if (fault == "source") f.Component.CodeModule.Source += "\r\n' changed during preparation";
                };
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveOtherHost(request, fault == "project path", fault == "null probe" ? null : f), fault);
                Assert.AreEqual(0, f.Attempts, fault);
            }
        }

        [TestMethod]
        public void EmptySourcesOptionalFormatsAndUnavailableProjectPathsHaveExactContracts()
        {
            var f = new Fixture(); f.Component.CodeModule.Source = ""; f.Component.CodeModule.LineCount = 0; f.Observation.Format = null;
            dynamic saved = f.Service.SaveOtherHost(f.Request(), false, f);
            Assert.IsTrue((bool)saved.Verified); Assert.AreEqual(1, f.Attempts);
            var helper = typeof(VbeProjectComponents).GetMethod("OtherHostProjectPath", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            f.Project.ComPathUnavailable = true; Assert.AreEqual("", helper.Invoke(null, new object[] { f.Project }));
            foreach (string path in new[] { null, "", "relative.docm" })
            {
                f = new Fixture(); var request = f.Request(); f.AfterInvocation = () => f.Observation.Path = path;
                dynamic result = f.Service.SaveOtherHost(request, false, f);
                Assert.IsTrue((bool)result.Uncertain); Assert.AreEqual(1, f.Attempts);
            }
        }

        [TestMethod]
        public void NativeApplicationResolutionUsesOwnedRotOrExactNativeOmFallbackForBothHosts()
        {
            var native = new VbeProjectComponents.NativeOtherHostProbe(); Assert.IsNull(native.HostKind);
            Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().Id, native.CurrentProcessId);
            Assert.ThrowsException<InvalidOperationException>(() => native.Application()); Assert.AreEqual(0u, native.ApplicationProcessId(new object()));
            foreach (string kind in new[] { "Word", "PowerPoint" }) foreach (string fault in new[] { "registered", "rot-com", "rot-binder", "rot-foreign", "om-error", "om-null", "om-com", "om-binder", "om-foreign", "om-absent" })
                using (var f = new VBAi.Tests.Infrastructure.NativeOtherHostsFixture())
                {
                    f.Kind = kind; foreach (long handle in new long[] { 20, 22, 23 }) f.Classes[handle] = kind == "Word" ? "_WwG" : "paneClassDC";
                    if (fault != "registered") f.Native.ReadActiveApplication = name => { if (fault == "rot-binder") throw new Microsoft.CSharp.RuntimeBinder.RuntimeBinderException("ROT binder"); if (fault == "rot-foreign") return new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.ApplicationContract(); throw new System.Runtime.InteropServices.COMException("ROT absent"); };
                    if (fault == "om-error") f.AccessibilityError = 1;
                    if (fault == "om-null") f.Accessible[20] = null;
                    if (fault == "om-com") f.Accessible[20] = new System.Runtime.InteropServices.COMException("NativeOM failed");
                    if (fault == "om-binder") f.Accessible[20] = new object();
                    if (fault == "om-foreign") f.Accessible[20] = new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.AutomationContract { Application = new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.ApplicationContract() };
                    if (fault == "om-absent") f.Children.Clear();
                    if (fault == "om-absent") Assert.ThrowsException<InvalidOperationException>(() => f.Native.Application());
                    else Assert.AreSame(f.Application, f.Native.Application(), kind + ":" + fault);
                    Assert.AreEqual(fault == "registered" || fault == "om-absent" ? 0 : fault.StartsWith("om-") ? 2 : 1, f.AccessibleCalls);
                }
        }

        [STATestMethod]
        public void NativeHostPidDocumentsStatesSaveAndFilesystemContractsAreFullyBounded()
        {
            var native = new VbeProjectComponents.NativeOtherHostProbe();
            using (var window = new System.Windows.Forms.Form()) { Assert.AreEqual((uint)System.Diagnostics.Process.GetCurrentProcess().Id, native.ReadOwner(window.Handle)); Assert.AreEqual(0u, native.ReadOwner(IntPtr.Zero)); }
            using (var f = new VBAi.Tests.Infrastructure.NativeOtherHostsFixture())
            {
                f.Application.Windows.Clear(); Assert.AreEqual(0u, f.Native.ApplicationProcessId(f.Application));
                f.Application.Windows.Add(new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.WindowContract { Hwnd = 1 });
                Assert.AreEqual((uint)f.Native.CurrentProcessId, f.Native.ApplicationProcessId(f.Application));
                f.Application.Windows.Add(new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.WindowContract { Hwnd = 2 }); Assert.AreEqual(0u, f.Native.ApplicationProcessId(f.Application));
                f.Owners[2] = 999999; Assert.AreEqual(0u, f.Native.ApplicationProcessId(f.Application));
                f.Application.Windows.Clear(); for (int i = 0; i < 1001; i++) f.Application.Windows.Add(new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.WindowContract { Hwnd = 1 });
                Assert.ThrowsException<InvalidOperationException>(() => f.Native.ApplicationProcessId(f.Application));
                foreach (string kind in new[] { "Word", "PowerPoint" })
                {
                    f.Kind = kind; var document = new VBAi.Tests.Infrastructure.NativeOtherHostsFixture.DocumentContract { VBProject = new object() };
                    f.Application.Documents.Clear(); f.Application.Presentations.Clear(); f.Application.Documents.Add(document); f.Application.Presentations.Add(document);
                    Assert.AreSame(document, f.Native.Documents(f.Application)[0]); Assert.AreSame(document.VBProject, f.Native.DocumentProject(document));
                    foreach (bool saved in new[] { false, true }) foreach (bool readOnly in new[] { false, true }) foreach (string path in new[] { null, " ", @"C:\Temp" })
                    {
                        document.Path = path; document.Saved = kind == "Word" ? (object)saved : saved ? -1 : 0; document.ReadOnly = kind == "Word" ? (object)readOnly : readOnly ? -1 : 0;
                        var state = f.Native.State(document); Assert.AreEqual(saved, state.Saved); Assert.AreEqual(readOnly, state.ReadOnly);
                        Assert.AreEqual(string.IsNullOrWhiteSpace(path) ? "" : document.FullName, state.Path); Assert.AreEqual(kind == "Word" ? (int?)13 : null, state.Format);
                    }
                    f.Native.Save(document, false, "ignored", 13); f.Native.Save(document, true, "destination", 27);
                    Assert.AreEqual(1, document.SaveCalls); Assert.AreEqual(kind == "Word" ? 1 : 0, document.WordSaveAsCalls); Assert.AreEqual(kind == "PowerPoint" ? 1 : 0, document.PowerPointSaveAsCalls);
                    Assert.AreEqual("destination", document.Destination); Assert.AreEqual(27, document.Format);
                    var catalogue = kind == "Word" ? f.Application.Documents : f.Application.Presentations; catalogue.Clear(); for (int i = 0; i < 1001; i++) catalogue.Add(document);
                    Assert.ThrowsException<InvalidOperationException>(() => f.Native.Documents(f.Application));
                }
                string directory = Path.Combine(Path.GetTempPath(), "VBAi-other-host-test-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
                try { string file = Path.Combine(directory, "proof.txt"); File.WriteAllText(file, "proof"); Assert.IsTrue(f.Native.DirectoryExists(directory)); Assert.IsTrue(f.Native.FileExists(file)); Assert.AreEqual(new FileInfo(file).Length, f.Native.FileLength(file)); Assert.IsFalse(f.Native.FileExists(file + ".missing")); }
                finally { Assert.IsTrue(Path.GetFullPath(directory).StartsWith(Path.GetFullPath(Path.GetTempPath()), StringComparison.OrdinalIgnoreCase)); Directory.Delete(directory, true); }
            }
        }

        [STATestMethod]
        public void NativeOtherHostProjectIdentityReleasesEveryAcquiredComReference()
        {
            var native = new VbeProjectComponents.NativeOtherHostProbe();
            object first = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true)), second = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
            object released = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true)); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(released);
            try
            {
                Assert.IsFalse(native.SameProject(null, first)); Assert.IsFalse(native.SameProject(first, null)); Assert.IsFalse(native.SameProject(new object(), first)); Assert.IsFalse(native.SameProject(first, new object()));
                Assert.IsTrue(native.SameProject(first, first)); Assert.IsFalse(native.SameProject(first, second));
                Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => native.SameProject(released, first)); Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => native.SameProject(first, released));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(first); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(second); }
        }
        /// <summary>Supported native extensions select exact Office constants and reject lossy formats.</summary>
        [TestMethod]
        public void HostRecognitionAndExactMacroFormatsNeverAcceptAnotherProcessOrLossyFormat()
        {
            Assert.AreEqual("Word", VbeProjectComponents.RecognizeOtherHost("WINWORD"));
            Assert.AreEqual("PowerPoint", VbeProjectComponents.RecognizeOtherHost("powerpnt"));
            foreach (string process in new[] { "EXCEL", "SLDWORKS", "WINWORD.exe", "Word", null }) Assert.IsNull(VbeProjectComponents.RecognizeOtherHost(process));
            Assert.AreEqual(13, VbeProjectComponents.OtherHostFormat("Word", "x.DOCM"));
            Assert.AreEqual(15, VbeProjectComponents.OtherHostFormat("Word", "x.dotm"));
            Assert.AreEqual(25, VbeProjectComponents.OtherHostFormat("PowerPoint", "x.pptm"));
            Assert.AreEqual(27, VbeProjectComponents.OtherHostFormat("PowerPoint", "x.potm"));
            Assert.AreEqual(29, VbeProjectComponents.OtherHostFormat("PowerPoint", "x.ppsm"));
            foreach (string extension in new[] { ".docx", ".pptx", ".pdf", ".xlsm", "" })
                foreach (string kind in new[] { "Word", "PowerPoint", "Other" }) Assert.ThrowsException<ArgumentException>(() => VbeProjectComponents.OtherHostFormat(kind, "x" + extension));
            var native = new VbeProjectComponents.NativeOtherHostProbe();
            Assert.IsFalse(native.SameProject(new object(), new object())); Assert.IsFalse(native.SameProject(null, null));
        }

        /// <summary>Le statut garde NOT_RUN et montre seulement les états d'un document associé par identité exacte.</summary>
        [TestMethod]
        public void PersistenceStatusRequiresOwnedProcessAndUniqueProjectIdentity()
        {
            var fixture = new Fixture(); dynamic status = fixture.Service.OtherHostPersistence("P", fixture);
            Assert.IsTrue((bool)status.HostAvailable); Assert.IsTrue((bool)status.IdentityVerified);
            Assert.AreEqual("NOT_RUN", (string)status.NativeQualification); Assert.AreEqual(0, fixture.Attempts);
            foreach (string scenario in new[] { "foreign", "none", "duplicate", "identity", "application" })
            {
                fixture = new Fixture();
                if (scenario == "foreign") fixture.Owner = 43;
                if (scenario == "none") fixture.Items.Clear();
                if (scenario == "duplicate") fixture.Items.Add(fixture);
                if (scenario == "identity") fixture.Identity = false;
                if (scenario == "application") fixture.Failure = "application";
                status = fixture.Service.OtherHostPersistence("P", fixture);
                Assert.IsFalse((bool)status.HostAvailable); Assert.AreEqual(0, fixture.Attempts);
            }
        }

        /// <summary>Save relit états, chemin, taille et code; SaveAs passe la bonne méthode et constante de chaque format macro.</summary>
        [TestMethod]
        public void ExistingSaveAndFirstSaveAsVerifyPathsFlagsBytesAndUnchangedLiveCode()
        {
            var fixture = new Fixture(); var request = fixture.Request();
            dynamic result = fixture.Service.SaveOtherHost(request, false, fixture);
            Assert.IsTrue((bool)result.Verified); Assert.IsTrue((bool)result.CodePreserved); Assert.AreEqual(1, fixture.Attempts);
            Assert.IsFalse((bool)result.PersistenceReopenVerified); Assert.AreEqual("NOT_RUN", (string)result.NativeQualification);
            foreach (string extension in new[] { ".docm", ".dotm", ".pptm", ".potm", ".ppsm" })
            {
                fixture = new Fixture { Kind = extension.StartsWith(".p", StringComparison.Ordinal) ? "PowerPoint" : "Word" };
                request = fixture.Request(true, extension); result = fixture.Service.SaveOtherHost(request, true, fixture);
                Assert.IsTrue((bool)result.Verified); Assert.IsTrue((bool)result.SaveAsInvoked); Assert.AreEqual(1, fixture.Attempts);
                Assert.AreEqual(VbeProjectComponents.OtherHostFormat(fixture.Kind, request.Path), fixture.ChosenFormat);
                Assert.AreEqual(fixture.Kind == "Word", (bool)result.NativeFileFormatVerified);
            }
        }

        /// <summary>Toute garde de la matrice antérieure à Save refuse la mutation.</summary>
        [TestMethod]
        public void VersionPathFormatReadonlyIdentityAndProcessGuardsPreventSaving()
        {
            foreach (string scenario in BeforeSaveFailures)
            {
                var fixture = new Fixture(); var request = fixture.Request();
                if (scenario == "unsupported") fixture.Kind = "Other";
                if (scenario == "foreign pid") fixture.Owner = 43;
                if (scenario == "missing document") fixture.Items.Clear();
                if (scenario == "duplicate document") fixture.Items.Add(fixture);
                if (scenario == "other identity") fixture.Identity = false;
                if (scenario == "identity read error" || scenario == "changed pid") fixture.Failure = scenario;
                if (scenario == "read only") fixture.Observation.ReadOnly = true;
                if (scenario == "stale version") request.ExpectedProjectVersion = "stale";
                if (scenario == "wrong host path") request.ExpectedHostPath = @"C:\fixture\Other.docm";
                if (scenario == "wrong project path")
                {
                    fixture.Kind = "PowerPoint"; fixture.Observation.Path = @"C:\fixture\Document.pptm"; fixture.Observation.Format = null;
                    request = fixture.Request(); fixture.Project.FileName = @"C:\fixture\Other.pptm";
                }
                if (scenario == "missing file") fixture.Exists = false;
                if (scenario == "native format") fixture.Observation.Format = 12;
                if (scenario == "protected") fixture.Project.Protection = 1;
                if (scenario == "runtime mode") fixture.Project.Mode = 1;
                if (scenario == "missing file") Assert.ThrowsException<FileNotFoundException>(() => fixture.Service.SaveOtherHost(request, false, fixture), scenario);
                else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveOtherHost(request, false, fixture), scenario);
                Assert.AreEqual(0, fixture.Attempts, scenario);
            }
        }

        /// <summary>La première sauvegarde refuse remplacement, dossiers, format sans macro, chemin relatif et document déjà sauvegardé.</summary>
        [TestMethod]
        public void FirstSaveAsRejectsEveryUnsafeDestinationBeforeInvocation()
        {
            foreach (string scenario in new[] { "exists", "directory", "parent", "relative", "format", "already saved", "expected path" })
            {
                var fixture = new Fixture(); var request = fixture.Request(scenario != "already saved");
                if (scenario == "exists") fixture.DestinationExists = true;
                if (scenario == "directory") fixture.DirectoryAtDestination = true;
                if (scenario == "parent") fixture.ParentExists = false;
                if (scenario == "relative") request.Path = "relative.docm";
                if (scenario == "format") request.Path = @"C:\fixture\New.docx";
                if (scenario == "expected path") request.ExpectedHostPath = @"C:\fixture\Old.docm";
                if (scenario == "exists" || scenario == "directory") Assert.ThrowsException<IOException>(() => fixture.Service.SaveOtherHost(request, true, fixture));
                else if (scenario == "parent") Assert.ThrowsException<DirectoryNotFoundException>(() => fixture.Service.SaveOtherHost(request, true, fixture));
                else if (scenario == "relative" || scenario == "format") Assert.ThrowsException<ArgumentException>(() => fixture.Service.SaveOtherHost(request, true, fixture));
                else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveOtherHost(request, true, fixture));
                Assert.AreEqual(0, fixture.Attempts);
            }
        }

        /// <summary>Après Save, chaque échec est déclaré incertain sans réessai, même si le fichier a déjà changé.</summary>
        [TestMethod]
        public void AfterInvocationFailuresRemainUncertainWithoutAutomaticRetry()
        {
            foreach (string scenario in AfterSaveFailures)
            {
                var fixture = new Fixture();
                if (scenario == "project path")
                {
                    fixture.Kind = "PowerPoint"; fixture.Observation.Path = @"C:\fixture\Document.pptm";
                    fixture.Project.FileName = fixture.Observation.Path; fixture.Observation.Format = null;
                }
                var request = fixture.Request();
                fixture.AfterInvocation = () =>
                {
                    if (scenario == "native error") fixture.Failure = "native error";
                    if (scenario == "host unsaved") fixture.Observation.Saved = false;
                    if (scenario == "project unsaved") fixture.Project.Saved = false;
                    if (scenario == "code changed") fixture.Component.CodeModule.Source += "\r\n' changed";
                    if (scenario == "host path") fixture.Observation.Path = @"C:\fixture\Other.docm";
                    if (scenario == "project path") fixture.Project.FileName = @"C:\fixture\Other.docm";
                    if (scenario == "format") fixture.Observation.Format = 12;
                    if (scenario == "empty file") fixture.Bytes = 0;
                    if (scenario == "missing file") fixture.Exists = false;
                    if (scenario == "identity") fixture.Identity = false;
                };
                dynamic result = fixture.Service.SaveOtherHost(request, false, fixture);
                Assert.IsFalse((bool)result.Verified, scenario); Assert.IsTrue((bool)result.Uncertain, scenario);
                Assert.IsTrue((bool)result.MutationInvoked, scenario); Assert.AreEqual(1, fixture.Attempts, scenario);
            }
        }
    }
}
