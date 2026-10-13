using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        /// <summary>Balances rejected window/application acquisitions and transfers only one matched application acquisition.</summary>
        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6), DataRow(7), DataRow(8), DataRow(9)]
        public void OwnedExcelResolutionBalancesTemporaryInterfaces(int scenario)
        {
            using (var scene = new SystemScene())
            {
                var own = scene.Add("Owned", "XLMAIN"); own.ProcessId = 42;
                var foreign = scene.Add("Other", "XLMAIN"); foreign.ProcessId = 43;
                scene.Add("Document", "EXCEL7", own).ProcessId = 42;
                if (scenario == 7) scene.Add("Second document", "EXCEL7", own).ProcessId = 42;
                int applicationReads = 0, accessibleReads = 0, fallbacks = 0;
                var released = new List<object>();
                var application = new ExcelResolutionObject
                {
                    ReadHandle = () => scenario == 5 ? throw new COMException("Closed application") :
                        scenario == 6 ? (object)"invalid HWND" : own.Handle.ToInt64()
                };
                var rejected = new ExcelResolutionObject { ReadHandle = () => foreign.Handle.ToInt64() };
                object returnedApplication = scenario == 1 || scenario == 7 ? rejected : scenario == 4 ? null : application;
                var window = new ExcelResolutionObject
                {
                    ReadApplication = () =>
                    {
                        applicationReads++;
                        if (scenario == 9) throw new COMException("Closed document");
                        return returnedApplication;
                    }
                };
                var nextWindow = new ExcelResolutionObject { ReadApplication = () => application };
                if (scenario == 8) { window.ReadHandle = () => own.Handle.ToInt64(); returnedApplication = window; }
                VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr handle, uint id, ref Guid iid, out object accessible) =>
                {
                    accessibleReads++;
                    accessible = scenario == 3 ? new object() : scenario == 7 && accessibleReads == 2 ? nextWindow : window;
                    return scenario == 2 ? -1 : 0;
                };
                object fallback = new object();
                Func<object> readFallback = () => { fallbacks++; return fallback; };
                if (scenario == 6)
                    Assert.ThrowsException<FormatException>(() => ExcelOwnedApplication.Resolve(42, readFallback, released.Add));
                else
                {
                    object expected = scenario == 0 || scenario == 7 ? application : scenario == 8 ? window : fallback;
                    Assert.AreSame(expected, ExcelOwnedApplication.Resolve(42, readFallback, released.Add));
                }
                if (scenario == 7)
                    CollectionAssert.AreEqual(new object[] { rejected, window, nextWindow }, released);
                else if (scenario == 1)
                    CollectionAssert.AreEqual(new object[] { rejected, window }, released);
                else if (scenario == 5 || scenario == 6)
                    CollectionAssert.AreEqual(new object[] { application, window }, released);
                else if (scenario == 3)
                { Assert.AreEqual(1, released.Count); Assert.AreNotSame(window, released[0]); }
                else CollectionAssert.AreEqual(new object[] { window }, released);
                Assert.AreEqual(scenario == 0 || scenario == 6 || scenario == 7 || scenario == 8 ? 0 : 1, fallbacks);
                Assert.AreEqual(scenario == 2 || scenario == 3 ? 0 : 1, applicationReads);
            }
        }

        /// <summary>An uncertain reference release stops resolution without another document probe or fallback.</summary>
        [TestMethod]
        public void OwnedExcelResolutionDoesNotSwallowReleaseFailureOrRetry()
        {
            using (var scene = new SystemScene())
            {
                var own = scene.Add("Owned", "XLMAIN"); own.ProcessId = 42;
                scene.Add("First", "EXCEL7", own).ProcessId = 42;
                scene.Add("Second", "EXCEL7", own).ProcessId = 42;
                int reads = 0, releases = 0, fallbacks = 0;
                VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr handle, uint id, ref Guid iid, out object accessible) =>
                { reads++; accessible = new object(); return -1; };
                Assert.ThrowsException<COMException>(() => ExcelOwnedApplication.Resolve(42,
                    () => { fallbacks++; return new object(); },
                    value => { releases++; throw new COMException("Release failed"); }));
                Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(0, fallbacks);
            }
        }

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
                    VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr handle, uint objectId, ref Guid iid, out object accessible) =>
                    {
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
