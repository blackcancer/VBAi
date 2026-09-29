using System;
using System.IO;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie les adaptateurs de sauvegarde par sondes injectables; qualification réelle Word/PowerPoint NOT_RUN.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed partial class VbeOtherHostPersistenceTests
    {
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
                f.BeforeSecondState = () => {
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
        /// <summary>Les extensions macro choisissent exactement les constantes Word/PowerPoint et refusent les autres formats.</summary>
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
                if (scenario == "wrong project path") fixture.Project.FileName = @"C:\fixture\Other.docm";
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
                var fixture = new Fixture(); var request = fixture.Request();
                fixture.AfterInvocation = () => {
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
