using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeProjectExcelHostTests
    {
        private static Fixture Create(string projectPath = null)
        {
            var project = new VbeProjectComponentsTests.FakeProject {
                FileName = projectPath ?? Path.Combine(Path.GetTempPath(),
                    "CodexVBE-Excel-host-" + Guid.NewGuid().ToString("N") + ".xlsm") };
            var vbe = new VbeProjectComponentsTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var workbook = new FakeWorkbook(project);
            workbook.FullName = project.FileName;
            workbook.Path = string.IsNullOrWhiteSpace(project.FileName) ? "" :
                Path.GetDirectoryName(project.FileName);
            var excel = new FakeExcel();
            excel.Workbooks.Add(workbook);
            var host = new FakeHost { Excel = excel };
            return new Fixture { Project = project, Workbook = workbook, Excel = excel, Host = host,
                Service = new VbeProjectComponents(vbe, new VbeForms(vbe), host) };
        }

        [TestMethod]
        public void StatusReadsTheMatchingWorkbookAndReportsUnavailableStates()
        {
            var fixture = Create();
            fixture.Workbook.VBASigned = true;
            fixture.Project.Saved = false;
            fixture.Workbook.Saved = false;
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            Assert.IsTrue((bool)signature.Available);
            Assert.AreEqual(true, (bool?)signature.Signed);
            dynamic persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsTrue((bool)persistence.HostAvailable);
            Assert.IsFalse((bool)persistence.ProjectSaved);
            Assert.AreEqual(false, (bool?)persistence.HostSaved);
            Assert.AreEqual(fixture.Project.FileName, (string)persistence.HostPath);

            fixture.Workbook.FullName = Path.Combine(Path.GetTempPath(), "different.xlsm");
            signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsFalse((bool)persistence.HostAvailable);
            StringAssert.Contains((string)persistence.Reason, "No workbook matches");
        }

        [TestMethod]
        public void StatusRejectsAnotherExcelProcessAndHandlesComLookupFailure()
        {
            var fixture = Create();
            fixture.Host.WindowOwner = 99;
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            dynamic persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsFalse((bool)persistence.HostAvailable);
            StringAssert.Contains((string)signature.Reason, "not this VBE host");
            fixture.Host.WindowOwner = (uint)fixture.Host.CurrentProcessId;
            fixture.Host.LookupFailure = new InvalidOperationException("ROT unavailable");
            signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            StringAssert.Contains((string)signature.Reason, "ROT unavailable");
            StringAssert.Contains((string)persistence.Reason, "ROT unavailable");
        }

        [TestMethod]
        public void SingleUnsavedWorkbookCanBeMatchedWithoutInventingAHostPath()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic status = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsTrue((bool)status.HostAvailable);
            Assert.AreEqual(false, (bool?)status.HostHasPath);
            Assert.IsNull((object)status.HostPath);
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            Assert.IsTrue((bool)signature.Available);
            fixture.Excel.Workbooks.Add(new FakeWorkbook(fixture.Project) { FullName = "Book2", Path = "" });
            status = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)status.HostAvailable);
            StringAssert.Contains((string)status.Reason, "no saved workbook path");
        }

        [TestMethod]
        public void SaveRequiresExactCurrentPathAndVerifiedSavedReadback()
        {
            var fixture = Create();
            fixture.Project.Saved = false;
            fixture.Workbook.Saved = false;
            dynamic projectState = fixture.Service.ProjectProperties(fixture.Project.Name);
            var request = new Request { Project = fixture.Project.Name,
                ExpectedProjectVersion = projectState.Version,
                ExpectedHostPath = fixture.Project.FileName };
            dynamic saved = fixture.Service.SaveHostDocument(request);
            Assert.IsTrue((bool)saved.SaveInvoked);
            Assert.IsFalse((bool)saved.HostSavedBefore);
            Assert.IsFalse((bool)saved.ProjectSavedBefore);
            Assert.IsTrue((bool)saved.HostSaved);
            Assert.IsTrue((bool)saved.ProjectSaved);
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);

            request.ExpectedHostPath = Path.Combine(Path.GetTempPath(), "different.xlsm");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SaveRefusesReadOnlyAndCancelledHostSave()
        {
            var fixture = Create();
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            var request = new Request { Project = fixture.Project.Name,
                ExpectedProjectVersion = state.Version, ExpectedHostPath = fixture.Project.FileName };
            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            Assert.AreEqual(0, fixture.Workbook.SaveAttempts);
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.CommitSave = false;
            fixture.Workbook.Saved = false;
            fixture.Project.Saved = false;
            request.ExpectedProjectVersion = ((dynamic)fixture.Service.ProjectProperties(fixture.Project.Name)).Version;
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            StringAssert.Contains(error.Message, "did not mark");
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SaveAsFirstTimeWritesMacroWorkbookAndVerifiesBothPaths()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            string path = Path.Combine(Path.GetTempPath(),
                "CodexVBE-SaveAs-" + Guid.NewGuid().ToString("N") + ".xlsm");
            try
            {
                dynamic result = fixture.Service.SaveHostDocumentAs(new Request {
                    Project = fixture.Project.Name, ExpectedProjectVersion = state.Version, Path = path });
                Assert.AreEqual(path, (string)result.HostPath);
                Assert.AreEqual(path, (string)result.ProjectPath);
                Assert.IsTrue((bool)result.SaveAsInvoked);
                Assert.IsTrue((long)result.Bytes > 0);
                Assert.AreEqual(52, fixture.Workbook.LastSaveAsFormat);
                Assert.AreEqual(1, fixture.Workbook.SaveAsAttempts);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        [TestMethod]
        public void SaveAsRefusesExistingDestinationReadOnlyWorkbookAndSecondSave()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            string path = Path.Combine(Path.GetTempPath(),
                "CodexVBE-SaveAs-" + Guid.NewGuid().ToString("N") + ".xlsm");
            var request = new Request { Project = fixture.Project.Name,
                ExpectedProjectVersion = state.Version, Path = path };
            request.Path = Path.ChangeExtension(path, ".xlsx");
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SaveHostDocumentAs(request));
            request.Path = path;
            File.WriteAllText(path, "existing");
            try { Assert.ThrowsException<IOException>(() => fixture.Service.SaveHostDocumentAs(request)); }
            finally { File.Delete(path); }
            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocumentAs(request));
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.Path = Path.GetTempPath();
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocumentAs(request));
            Assert.AreEqual(0, fixture.Workbook.SaveAsAttempts);
        }

        [TestMethod]
        public void SignedWorkbookPersistenceRequiresSignatureBeforeAndAfterSave()
        {
            var fixture = Create();
            fixture.Workbook.VBASigned = true;
            dynamic result = fixture.Service.PersistExcelSignature(fixture.Project.Name);
            Assert.IsTrue((bool)result.Available);
            Assert.IsTrue((bool)result.Saved);
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);

            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PersistExcelSignature(fixture.Project.Name));
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.VBASigned = false;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PersistExcelSignature(fixture.Project.Name));
            fixture.Workbook.VBASigned = true;
            fixture.Workbook.DropSignatureOnSave = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                fixture.Service.PersistExcelSignature(fixture.Project.Name));
            StringAssert.Contains(error.Message, "no longer reports");
            Assert.AreEqual(2, fixture.Workbook.SaveAttempts);
        }

        private sealed class Fixture
        {
            public VbeProjectComponentsTests.FakeProject Project;
            public FakeWorkbook Workbook;
            public FakeExcel Excel;
            public FakeHost Host;
            public VbeProjectComponents Service;
        }

        public sealed class FakeHost : VbeProjectComponents.IExcelHostProbe
        {
            public bool IsExcel { get; set; } = true;
            public int CurrentProcessId { get; set; } = 42;
            public uint WindowOwner { get; set; } = 42;
            public FakeExcel Excel { get; set; }
            public Exception LookupFailure { get; set; }
            public object ExcelApplication()
            {
                if (LookupFailure != null) throw LookupFailure;
                return Excel;
            }
            public uint WindowProcessId(IntPtr window) { return WindowOwner; }
        }

        public sealed class FakeExcel
        {
            public int Hwnd { get; set; } = 100;
            public FakeWorkbooks Workbooks { get; } = new FakeWorkbooks();
        }

        public sealed class FakeWorkbooks : IEnumerable<FakeWorkbook>
        {
            private readonly List<FakeWorkbook> items = new List<FakeWorkbook>();
            public int Count => items.Count;
            public void Add(FakeWorkbook workbook) { items.Add(workbook); }
            public FakeWorkbook Item(int index) { return items[index - 1]; }
            public IEnumerator<FakeWorkbook> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeWorkbook
        {
            private readonly VbeProjectComponentsTests.FakeProject project;
            public FakeWorkbook(VbeProjectComponentsTests.FakeProject project) { this.project = project; }
            public string FullName { get; set; }
            public string Path { get; set; }
            public bool Saved { get; set; } = true;
            public bool ReadOnly { get; set; }
            public bool VBASigned { get; set; }
            public bool CommitSave { get; set; } = true;
            public bool DropSignatureOnSave { get; set; }
            public int SaveAttempts { get; private set; }
            public int SaveAsAttempts { get; private set; }
            public int LastSaveAsFormat { get; private set; }
            public void Save()
            {
                SaveAttempts++;
                if (CommitSave) { Saved = true; project.Saved = true; }
                if (DropSignatureOnSave) VBASigned = false;
            }
            public void SaveAs(string path, int format)
            {
                SaveAsAttempts++;
                LastSaveAsFormat = format;
                File.WriteAllText(path, "macro workbook test fixture");
                FullName = path;
                Path = System.IO.Path.GetDirectoryName(path);
                project.FileName = path;
                Saved = true;
                project.Saved = true;
            }
        }
    }
}
