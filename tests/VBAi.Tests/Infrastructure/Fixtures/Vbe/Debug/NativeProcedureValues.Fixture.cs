using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Real owned HWND and IUnknown identities, with a managed application dispatch contract; never Excel.</summary>
    public sealed class NativeProcedureValuesFixture : IDisposable
    {
        internal readonly Form Owner = new Form();
        internal readonly VbeDebug.NativeProcedureValuesHost Host = new VbeDebug.NativeProcedureValuesHost();
        internal readonly ApplicationContract Application = new ApplicationContract();
        internal readonly WorkbookContract Workbook;
        internal readonly object Project = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
        internal readonly object OtherProject = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
        internal const string Path = @"C:\Temp\Owner's.xlsm";
        internal NativeProcedureValuesFixture()
        {
            Application.Hwnd = Owner.Handle.ToInt64(); Workbook = new WorkbookContract { VBProject = Project, FullName = Path, Path = @"C:\Temp" }; Application.Workbooks.Add(Workbook);
            Host.ReadProcessName = () => "EXCEL";
            Host.ResolveApplication = (pid, read) => { Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().Id, pid); return read(); };
            Host.ReadActiveApplication = name => { Assert.AreEqual("Excel.Application", name); return Application; };
        }
        internal object Resolve() => Host.ResolveTarget(Project, Path);
        public void Dispose() { Owner.Dispose(); Marshal.FinalReleaseComObject(Project); Marshal.FinalReleaseComObject(OtherProject); }
        public sealed class WorkbookContract { public object VBProject { get; set; } public string FullName { get; set; } public string Path { get; set; } }
        public sealed class ApplicationContract
        {
            public long Hwnd { get; set; }
            public List<WorkbookContract> Workbooks { get; } = new List<WorkbookContract>();
            public string Macro; public object[] Arguments; public int Invocations;
            public object Run(string macro, object first = null, object second = null)
            { Macro = macro; Arguments = new[] { first, second }; Invocations++; return "Native dispatch contract return"; }
        }
    }
}
