using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestWordValuesHostTests
    {
        [TestMethod]
        public void DefaultRunUsesDocumentQualifiedNameAndZeroOrTwoPositionalArguments()
        {
            using (var fixture = new Fixture())
            {
                object native = new object(); fixture.Application.Result = native;
                var target = fixture.Resolve();
                Assert.AreSame(native, fixture.Host.Invoke(target, "Support", "Reset", new object[0]));
                Assert.AreEqual("'Original.docm'!Support.Reset", fixture.Application.LastMacro);
                Assert.IsNull(fixture.Application.First);
                Assert.AreSame(native, fixture.Host.Invoke(target, "Support", "Run", new object[] { "Tests", "Alpha" }));
                Assert.AreEqual("Tests", fixture.Application.First);
                Assert.AreEqual("Alpha", fixture.Application.Second);
                Assert.AreEqual(2, fixture.Application.RunCalls);
                Assert.AreEqual(2, fixture.Source.ActivateCalls);
            }
        }

        [TestMethod]
        public void WrongProcessPidProjectAndStalePathNeverActivateOrRun()
        {
            using (var fixture = new Fixture())
            {
                fixture.Host.ReadProcessName = () => "EXCEL";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(0, fixture.Reads);
                fixture.Host.ReadProcessName = () => "WINWORD";
                fixture.Host.ReadWindowOwner = hwnd => 999;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                fixture.Host.ReadWindowOwner = hwnd => 123;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(new object(), fixture.Source.FullName));
                var target = fixture.Resolve();
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Renamed.docm");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Source.ActivateCalls + fixture.Application.RunCalls);
            }
        }

        [TestMethod]
        public void DuplicateFilenamesUnsafeQuotingAndUnsupportedArgumentsRefuseBeforeActivation()
        {
            using (var fixture = new Fixture())
            {
                var target = fixture.Resolve();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[] { 1 }));
                fixture.Application.Documents.Add(new Document { FullName = Path.Combine(fixture.Folder, "Other", "Original.docm"), Application = fixture.Application });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                fixture.Application.Documents.RemoveAt(1);
                fixture.Source.FullName = Path.Combine(fixture.Folder, "User's.docm");
                target = fixture.Resolve();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Source.ActivateCalls + fixture.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ActivationContextFailureIsUncertainWithoutFallbackOrRetry()
        {
            foreach (string fault in new[] { "path", "context", "pid", "name" })
            using (var fixture = new Fixture())
            {
                var target = fixture.Resolve();
                fixture.Source.OnActivate = () => {
                    if (fault == "path") fixture.Source.FullName = Path.Combine(fixture.Folder, "Changed.docm");
                    if (fault == "context") fixture.Application.ActiveDocument = new Document();
                    if (fault == "pid") fixture.Host.ReadWindowOwner = hwnd => 999;
                    if (fault == "name") fixture.Application.Documents.Add(new Document { FullName = Path.Combine(fixture.Folder, "Other", "Original.docm") });
                };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, fixture.Source.ActivateCalls, fault);
                Assert.AreEqual(0, fixture.Application.RunCalls, fault);
            }
        }

        [TestMethod]
        public void NativeMacroFailureIsUncertainAndAttemptedExactlyOnce()
        {
            using (var fixture = new Fixture())
            {
                int attempts = 0;
                fixture.Host.RunProcedure = (application, macro, arguments) => { attempts++; throw new InvalidOperationException("Native completion unavailable"); };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Host.Invoke(fixture.Resolve(), "Support", "Run", new object[] { "Tests", "Alpha" }));
                Assert.IsTrue(error.Uncertain);
                Assert.AreEqual(1, attempts);
            }
        }

        [TestMethod]
        public void WrongThreadAndDriveRelativePathsCannotReadWordCom()
        {
            using (var fixture = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(fixture.Source.VBProject, @"C:Original.docm"));
                Exception error = null;
                var thread = new Thread(() => { try { fixture.Resolve(); } catch (Exception caught) { error = caught; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
                Assert.AreEqual(0, fixture.Reads);
            }
        }

        [TestMethod]
        public void MissingOrForeignTargetsAndInvalidIdentifiersRefuseBeforeActivation()
        {
            using (var f = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ResolveTarget(null, f.Source.FullName));
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(null, "Support", "Run", null));
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(new VbaTestWordValuesHost.OwnedTarget { Owner = new VbaTestWordValuesHost() }, "Support", "Run", null));
                var target = f.Resolve();
                foreach (var pair in new[] { new[] { (string)null, "Run" }, new[] { "bad!", "Run" }, new[] { "Support", (string)null }, new[] { "Support", "bad!" } })
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, pair[0], pair[1], null));
                f.Host.Invoke(target, "Support", "Run", null);
                Assert.AreEqual(1, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ChangedApplicationOrDocumentIdentityAndUnboundedCollectionAreRejected()
        {
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Host.ReadActiveApplication = _ => null;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadActiveApplication = _ => new Application();
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadActiveApplication = _ => f.Application;
                var replacement = new Document { FullName = f.Source.FullName, VBProject = f.Source.VBProject, Application = f.Application };
                f.Application.Documents[0] = replacement;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Application.Documents.Add(f.Source);
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.Documents.Clear(); f.Application.Documents.Add(f.Source);
                for (int i = 0; i < 1000; i++) f.Application.Documents.Add(new Document());
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.AreEqual(0, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void NameAndPathValidationRejectUnsafeSavedMacroContexts()
        {
            using (var f = new Fixture())
            {
                f.Source.PathOverride = " ";
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Source.PathOverride = null;
                var target = f.Resolve();
                f.Source.NameOverride = "Other.docm";
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                Assert.AreEqual(0, f.Source.ActivateCalls);
            }
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(null, @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"C:\A.docm", "A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"\A.docm", @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"C:A.docm", @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(" ", null));
        }

        [TestMethod]
        public void DefaultProcessAndWindowReadersObserveOnlyTheTestProcess()
        {
            var host = new VbaTestWordValuesHost();
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
            {
                Assert.AreEqual(process.ProcessName, host.ReadProcessName());
                Assert.AreEqual(process.Id, host.ReadProcessId());
            }
            Assert.AreEqual(0u, host.ReadWindowOwner(IntPtr.Zero));
            Assert.IsFalse(host.SameIdentity(new object(), new object()));
        }

        [TestMethod]
        public void ActiveWordWindowHandleIsRequiredBeforeDocumentsCanBeInspected()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(typeof(Application).GetProperty("Hwnd"), "The Word application model must not expose the nonexistent application HWND.");
                Assert.AreEqual(new IntPtr(99), VbaTestWordValuesHost.ReadApplicationWindow(f.Application));
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestWordValuesHost.ReadApplicationWindow(null));
                f.Application.ActiveWindow = null;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.ActiveWindow = new Window();
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.ActiveWindow.Hwnd = 99;
                f.Host.ReadWindowOwner = _ => 0;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Host.ReadWindowOwner = _ => 999;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.AreEqual(0, f.Source.ActivateCalls + f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ActiveWordWindowOwnershipIsRecheckedAfterDocumentActivation()
        {
            foreach (string fault in new[] { "missing", "zero", "foreign" })
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Source.OnActivate = () =>
                {
                    if (fault == "missing") f.Application.ActiveWindow = null;
                    if (fault == "zero") f.Application.ActiveWindow.Hwnd = 0;
                    if (fault == "foreign") f.Host.ReadWindowOwner = _ => 999;
                };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, f.Source.ActivateCalls, fault);
                Assert.AreEqual(0, f.Application.RunCalls, fault);
            }
        }
        internal sealed class Fixture : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-Word-" + Guid.NewGuid().ToString("N"));
            internal readonly Application Application = new Application();
            internal readonly Document Source;
            internal readonly VbaTestWordValuesHost Host = new VbaTestWordValuesHost();
            internal int Reads;
            internal Fixture(string extension = ".docm")
            {
                Directory.CreateDirectory(Folder);
                Source = new Document { FullName = Path.Combine(Folder, "Original" + extension), Application = Application };
                File.WriteAllText(Source.FullName, "Saved Word fixture bytes");
                Application.Documents.Add(Source);
                Host.ReadProcessName = () => "WINWORD"; Host.ReadProcessId = () => 123;
                Host.ReadWindowOwner = hwnd => { Assert.AreEqual(new IntPtr(99), hwnd); return 123; };
                Host.ReadActiveApplication = progId => { Reads++; Assert.AreEqual("Word.Application", progId); return Application; };
            }
            internal object Resolve() => Host.ResolveTarget(Source.VBProject, Source.FullName);
            public void Dispose() { if (Directory.Exists(Folder)) Directory.Delete(Folder, true); }
        }

        public sealed class Application
        {
            public Window ActiveWindow { get; set; } = new Window { Hwnd = 99 };
            public Documents Documents { get; } = new Documents();
            public Document ActiveDocument { get; set; }
            public object Result = new object();
            public string LastMacro;
            public object First, Second;
            public int RunCalls;
            public object Run(string macro, object first = null, object second = null)
            { RunCalls++; LastMacro = macro; First = first; Second = second; return Result; }
        }
        public sealed class Window
        {
            public int Hwnd { get; set; }
        }
        public sealed class Documents : List<Document>
        {
            public int OpenCalls;
            public bool LastReadOnly, LastRecent, LastVisible, LastConversions, LastRevert, LastRepair;
            public Func<string, Document> OnOpen;
            public Document Open(string FileName, bool ConfirmConversions, bool ReadOnly, bool AddToRecentFiles, bool Revert, bool Visible, bool OpenAndRepair)
            {
                OpenCalls++; LastReadOnly = ReadOnly; LastRecent = AddToRecentFiles; LastVisible = Visible;
                LastConversions = ConfirmConversions; LastRevert = Revert; LastRepair = OpenAndRepair;
                var document = OnOpen(FileName);
                if (document != null && !Contains(document)) Add(document);
                return document;
            }
        }
        public sealed class Document
        {
            public object VBProject { get; set; } = new object();
            public string FullName { get; set; }
            public object Saved { get; set; } = true;
            public string NameOverride, PathOverride;
            public string Name => NameOverride ?? System.IO.Path.GetFileName(FullName);
            public string Path => PathOverride ?? System.IO.Path.GetDirectoryName(FullName);
            public Application Application;
            public Action OnActivate;
            public bool CancelClose;
            public int ActivateCalls, CloseCalls, LastSaveChanges, SaveCalls, SaveAsCalls;
            public void Activate() { ActivateCalls++; Application.ActiveDocument = this; OnActivate?.Invoke(); }
            public void Close(int SaveChanges) { CloseCalls++; LastSaveChanges = SaveChanges; if (!CancelClose) Application.Documents.Remove(this); }
        }
    }
}
