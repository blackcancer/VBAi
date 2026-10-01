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
            public int Hwnd => 99;
            public Documents Documents { get; } = new Documents();
            public Document ActiveDocument { get; set; }
            public object Result = new object();
            public string LastMacro;
            public object First, Second;
            public int RunCalls;
            public object Run(string macro, object first = null, object second = null)
            { RunCalls++; LastMacro = macro; First = first; Second = second; return Result; }
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
            public bool Saved { get; set; } = true;
            public string Name => System.IO.Path.GetFileName(FullName);
            public string Path => System.IO.Path.GetDirectoryName(FullName);
            public Application Application;
            public Action OnActivate;
            public bool CancelClose;
            public int ActivateCalls, CloseCalls, LastSaveChanges, SaveCalls, SaveAsCalls;
            public void Activate() { ActivateCalls++; Application.ActiveDocument = this; OnActivate?.Invoke(); }
            public void Close(int SaveChanges) { CloseCalls++; LastSaveChanges = SaveChanges; if (!CancelClose) Application.Documents.Remove(this); }
        }
    }
}
