using System;
using System.Runtime.InteropServices;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void OwnedExcelResolutionSkipsOtherProcessesAndKeepsFallbackForUnavailableDocuments()
        {
            for (int scenario = 0; scenario < 7; scenario++)
            using (var scene = new SystemScene())
            {
                var other = scene.Add("Other", "XLMAIN"); other.ProcessId = 43;
                scene.Add("Other document", "EXCEL7", other).ProcessId = 43;
                var own = scene.Add("Owned", "XLMAIN"); own.ProcessId = 42;
                scene.Add("Toolbar", "MsoCommandBar", own).ProcessId = 42;
                scene.Add("Foreign child", "EXCEL7", own).ProcessId = 43;
                var first = scene.Add("Document", "EXCEL7", own); first.ProcessId = 42;
                var second = scene.Add("Second document", "EXCEL7", own); second.ProcessId = 42;
                var application = new ExcelApplicationObject { Hwnd = own.Handle.ToInt64() };
                int reads = 0, fallbacks = 0;
                VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr handle, uint objectId, ref Guid iid, out object accessible) => {
                    Assert.AreEqual(0xFFFFFFF0u, objectId); Assert.AreEqual(new Guid("00020400-0000-0000-C000-000000000046"), iid);
                    reads++; accessible = null;
                    if (scenario == 0) return -1;
                    if (scenario == 1) return 0;
                    if (scenario == 2) throw new COMException("closed document");
                    if (scenario == 3) { accessible = new object(); return 0; }
                    if (scenario == 4) { accessible = new ExcelDocumentObject { Application = new ExcelApplicationObject { Hwnd = other.Handle.ToInt64() } }; return 0; }
                    if (scenario == 5 && reads == 1) return -1;
                    accessible = new ExcelDocumentObject { Application = application }; return 0;
                };
                object fallback = new object();
                object result = ExcelOwnedApplication.Resolve(42, () => { fallbacks++; return fallback; });
                Assert.AreSame(scenario >= 5 ? application : fallback, result);
                Assert.AreEqual(scenario >= 5 ? 0 : 1, fallbacks); Assert.IsTrue(reads > 0);
            }
            using (var scene = new SystemScene())
            {
                object fallback = new object(); Assert.AreSame(fallback, ExcelOwnedApplication.Resolve(42, () => fallback));
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedApplication.Resolve(42, () => throw new InvalidOperationException("ROT unavailable")));
            }
        }
    }
}
