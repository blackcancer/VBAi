namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    public sealed partial class VbeProjectExcelHostTests
    {
        private static Fixture Create(string projectPath = null)
        {
            var project = new VbeProjectComponentsTests.FakeProject
            {
                FileName = projectPath ?? Path.Combine(Path.GetTempPath(), "VBAi-Excel-host-" + Guid.NewGuid().ToString("N") + ".xlsm")
            };
            var vbe = new VbeProjectComponentsTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var workbook = new FakeWorkbook(project);
            workbook.FullName = project.FileName;
            workbook.Path = string.IsNullOrWhiteSpace(project.FileName) ? "" : Path.GetDirectoryName(project.FileName);
            var excel = new FakeExcel();
            excel.Workbooks.Add(workbook);
            var host = new FakeHost
            {
                Excel = excel
            };
            return new Fixture
            {
                Project = project,
                Workbook = workbook,
                Excel = excel,
                Host = host,
                Vbe = vbe,
                Service = new VbeProjectComponents(vbe, new VbeForms(vbe), host)
            };
        }

        private sealed class Fixture
        {
            public VbeProjectComponentsTests.FakeProject Project;
            public FakeWorkbook Workbook;
            public FakeExcel Excel;
            public FakeHost Host;
            public VbeProjectComponents Service;
            public VbeProjectComponentsTests.FakeVbe Vbe;
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
                if (LookupFailure != null)
                    throw LookupFailure;
                return Excel;
            }

            public uint WindowProcessId(IntPtr window)
            {
                return WindowOwner;
            }
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

            public void Add(FakeWorkbook workbook)
            {
                items.Add(workbook);
            }
            public void Clear() { items.Clear(); }

            public FakeWorkbook Item(int index)
            {
                return items[index - 1];
            }

            public IEnumerator<FakeWorkbook> GetEnumerator()
            {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeWorkbook
        {
            private readonly VbeProjectComponentsTests.FakeProject project;
            public FakeWorkbook(VbeProjectComponentsTests.FakeProject project)
            {
                this.project = project;
            }

            public string FullName { get; set; }
            public string Path { get; set; }
            public int FileFormat { get; set; } = 52;
            public bool Saved { get; set; } = true;
            public bool ReadOnly { get; set; }
            public bool VBASigned { get; set; }
            public bool CommitSave { get; set; } = true;
            public bool DropSignatureOnSave { get; set; }
            public int SaveAttempts { get; private set; }
            public int SaveAsAttempts { get; private set; }
            public int LastSaveAsFormat { get; private set; }
            public Action AfterSaveAs { get; set; }
            public Action AfterSave { get; set; }

            public void Save()
            {
                SaveAttempts++;
                if (CommitSave)
                {
                    Saved = true;
                    project.Saved = true;
                }

                if (DropSignatureOnSave)
                    VBASigned = false;
                AfterSave?.Invoke();
            }

            public void SaveAs(string path, int format)
            {
                SaveAsAttempts++;
                LastSaveAsFormat = format;
                FileFormat = format;
                File.WriteAllText(path, "macro workbook test fixture");
                FullName = path;
                Path = System.IO.Path.GetDirectoryName(path);
                project.FileName = path;
                Saved = true;
                project.Saved = true;
                AfterSaveAs?.Invoke();
            }
        }
    }
}
