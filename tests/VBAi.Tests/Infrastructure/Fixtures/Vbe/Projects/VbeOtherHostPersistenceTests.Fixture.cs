using System;
using System.Collections.Generic;
using System.IO;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeOtherHostPersistenceTests
    {
        /// <summary>Matrice préalable: identité COM/PID, lecture seule, versions, chemins, formats, destination et incertitude après sauvegarde.</summary>
        private static readonly string[] BeforeSaveFailures = { "unsupported", "foreign pid", "missing document", "duplicate document", "other identity", "identity read error",
            "read only", "stale version", "wrong host path", "wrong project path", "missing file", "native format", "protected", "runtime mode", "changed pid" };
        /// <summary>Après invocation: erreur native, annulation, code modifié, mauvais chemins/format, fichier vide/absent et identité changée.</summary>
        private static readonly string[] AfterSaveFailures = { "native error", "host unsaved", "project unsaved", "code changed", "host path", "project path", "format", "empty file", "missing file", "identity" };

        /// <summary>Projet public pour les appels dynamiques de VBIDE simulé.</summary>
        public sealed class OtherProject : VbeProjectComponentsTests.FakeProject
        {
            public int Protection { get; set; }
            public bool ComPathUnavailable;
            public override string FileName { get { if (ComPathUnavailable) throw new System.Runtime.InteropServices.COMException("Unsaved project"); return base.FileName; } set { base.FileName = value; } }
        }

        /// <summary>Sonde complète de sauvegarde ne touchant aucun vrai document Office ni fichier.</summary>
        private sealed class Fixture : VbeProjectComponents.IOtherHostProbe
        {
            internal readonly OtherProject Project = new OtherProject { Name = "P", FileName = @"C:\fixture\Document.docm", Saved = true };
            internal readonly VbeProjectComponentsTests.FakeComponent Component = new VbeProjectComponentsTests.FakeComponent("Code", 1);
            internal readonly List<object> Items = new List<object>();
            internal readonly VbeProjectComponents Service;
            internal VbeProjectComponents.OtherHostDocumentState Observation = new VbeProjectComponents.OtherHostDocumentState { Path = @"C:\fixture\Document.docm", Format = 13, Saved = true };
            internal string Kind = "Word", ActualPath, Failure;
            internal uint Owner = 42;
            internal bool Identity = true, Exists = true, ParentExists = true, DestinationExists, DirectoryAtDestination;
            internal int Calls, Attempts, ChosenFormat;
            internal long Bytes = 10;
            internal Action AfterInvocation;
            internal Action BeforeSecondState;
            internal int ProcessId = 42, StateCalls, IdentityCalls;
            internal bool NullDocuments, ChangeIdentity;
            public string HostKind => Kind;
            public int CurrentProcessId => ProcessId;
            internal Fixture()
            {
                Project.VBComponents.Add(Component); Items.Add(this);
                var vbe = new VbeProjectComponentsTests.FakeVbe(); vbe.VBProjects.Add(Project);
                Service = new VbeProjectComponents(vbe, new VbeForms(vbe));
            }
            internal Request Request(bool unsaved = false, string extension = ".docm")
            {
                if (unsaved) { Project.FileName = ""; Observation.Path = ""; Exists = false; }
                dynamic snapshot = Service.ProjectProperties("P");
                return new Request { Project = "P", ExpectedProjectVersion = snapshot.Version,
                    ExpectedHostPath = unsaved ? "" : Observation.Path, Path = @"C:\fixture\New" + extension };
            }
            public object Application() { if (Failure == "application") throw new InvalidOperationException("application unreadable"); return this; }
            public uint ApplicationProcessId(object app) { Calls++; return Failure == "changed pid" && Calls > 1 ? 43 : Owner; }
            public IList<object> Documents(object app) => NullDocuments ? null : Items;
            public object DocumentProject(object document) { if (Failure == "identity read error") throw new InvalidOperationException("identity unreadable"); return Project; }
            public bool SameProject(object first, object second) { IdentityCalls++; return Identity && !(ChangeIdentity && IdentityCalls > 1) && ReferenceEquals(first, second); }
            public VbeProjectComponents.OtherHostDocumentState State(object document) { StateCalls++; if (StateCalls == 2) BeforeSecondState?.Invoke(); return new VbeProjectComponents.OtherHostDocumentState { Path = Observation.Path, Format = Observation.Format, Saved = Observation.Saved, ReadOnly = Observation.ReadOnly }; }
            public void Save(object document, bool saveAs, string path, int format)
            {
                Attempts++; ActualPath = path; ChosenFormat = format;
                Project.FileName = path; Observation.Path = path; Project.Saved = true; Observation.Saved = true;
                Observation.Format = Kind == "Word" ? (int?)format : null; Exists = true;
                AfterInvocation?.Invoke();
                if (Failure == "native error") throw new InvalidOperationException("Native save failed after invocation");
            }
            public bool FileExists(string path) => path == ActualPath || path == Observation.Path ? Exists : DestinationExists;
            public bool DirectoryExists(string path) => Path.GetExtension(path) != "" ? DirectoryAtDestination : ParentExists;
            public long FileLength(string path) => Bytes;
        }
    }
}
