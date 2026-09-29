namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using VBAi;
    public sealed partial class VbeProjectComponentsTests
    {
        public sealed class FaultedStandaloneProject : FakeProject
        {
            public int Type{get;set;}=101;
            public Action<string> Saving;
            public int Saves;
            public void SaveAs(string destination){Saves++;if(Saving!=null)Saving(destination);else{File.WriteAllText(destination,"native fixture");FileName=destination;Saved=true;}}
        }
    }
    public sealed partial class VbeProjectExcelHostTests
    {
        private sealed class RenameWorkflowFixture
        {
            internal readonly RenameFaultProject Project=new RenameFaultProject();
            internal readonly VbeProjectComponentsTests.FakeVbe Vbe=new VbeProjectComponentsTests.FakeVbe();
            internal readonly RenameWorkbook Workbook;
            internal readonly RenameHost Host=new RenameHost();
            internal readonly VbeProjectComponents Service;
            internal RenameWorkflowFixture()
            {
                Project.FileName=Path.Combine(Path.GetTempPath(),"rename.xlsm");Project.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("Module1",1));Vbe.VBProjects.Add(Project);
                Workbook=new RenameWorkbook{FullName=Project.FileName};Host.Excel.Workbooks.Add(Workbook);Service=new VbeProjectComponents(Vbe,new VbeForms(Vbe),Host);
            }
            internal Request Request(string name="Renamed"){dynamic metadata=Service.ProjectProperties(Project.FileName);return new Request{Project=Project.FileName,Property="Name",Value=name,ExpectedProjectVersion=metadata.Version};}
        }
        public sealed class RenameHost : VbeProjectComponents.IExcelHostProbe
        {
            public bool IsExcel{get;set;}=true;
            public int CurrentProcessId=>42;
            public uint WindowProcessId(IntPtr handle)=>42;
            public readonly RenameExcel Excel=new RenameExcel();
            public object ExcelApplication()=>Excel;
        }
        public sealed class RenameExcel{public int Hwnd=>10;public List<RenameWorkbook> Workbooks{get;}=new List<RenameWorkbook>();}
        public sealed class RenameWorkbook
        {
            public string FullName{get;set;}
            public Action ReadingReadOnly;
            public bool ReadOnly{get{ReadingReadOnly?.Invoke();return false;}}
        }
        public sealed class RenameFaultProject : VbeProjectComponentsTests.FakeProject
        {
            public Action OnRename;
            public Func<string,string> ReadPath;
            public bool IgnoreName,FailProtection;
            public int ProtectionValue;
            public int Protection{get{if(FailProtection)throw new InvalidOperationException("protection unavailable");return ProtectionValue;}}
            public override string Name{get=>base.Name;set{if(!IgnoreName)base.Name=value;OnRename?.Invoke();}}
            public override string FileName{get{var path=base.FileName;return ReadPath==null?path:ReadPath(path);}set=>base.FileName=value;}
        }
    }
}
