using System;
using System.IO;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie les adaptateurs de sauvegarde par sondes injectables; qualification réelle Word/PowerPoint NOT_RUN.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed partial class VbeOtherHostPersistenceTests
    {
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
