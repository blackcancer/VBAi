using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class IsolatedTestDesktopMainTests
    {
        [DataTestMethod]
        [DataRow("WinSta0", "Default", "Default", true)]
        [DataRow("winsta0", "default", "DEFAULT", true)]
        [DataRow(null, "Default", "Default", false)]
        [DataRow("Service", "Default", "Default", false)]
        [DataRow("WinSta0", null, "Default", false)]
        [DataRow("WinSta0", "Private", "Default", false)]
        [DataRow("WinSta0", "Default", null, false)]
        [DataRow("WinSta0", "Default", "Winlogon", false)]
        [DataRow("WinSta0", "Default ", "Default", false)]
        public void InteractivePlacementRequiresActualStationCurrentAndInput(string station, string current, string input, bool valid)
        {
            if (valid) IsolatedTestDesktop.RequireMainObserved(station, current, input);
            else Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireMainObserved(station, current, input));
        }

        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(6)]
        public void MainLaunchChecksBeforeOneCreateAndNeverReplaysAnUncertainFailure(int failure)
        {
            var log = new List<string>(); var error = new InvalidOperationException("same native/guard failure");
            object originalHandleOwner = new object();
            Func<object> run = () => IsolatedTestDesktop.LaunchMainOnce(failure == 1 ? "relative.exe" : @"C:\Owned\WINWORD.EXE",
                new[] { "/a", @"C:\Owned\seed.docx" }, failure == 2 ? "relative" : @"C:\Owned",
                () => { log.Add("guard"); if (failure == 5) throw error; },
                path => { log.Add("file"); return failure != 3; }, path => { log.Add("directory"); return failure != 4; },
                (path, args, directory) => { log.Add("create"); CollectionAssert.AreEqual(new[] { "/a", @"C:\Owned\seed.docx" }, args);
                    Assert.AreEqual(@"C:\Owned", directory); if (failure == 6) throw error; return originalHandleOwner; });
            if (failure == 0) { Assert.AreSame(originalHandleOwner, run()); CollectionAssert.AreEqual(new[] { "file", "directory", "guard", "create" }, log); }
            else if (failure < 5) { Assert.ThrowsException<ArgumentException>(() => run()); Assert.IsFalse(log.Contains("guard")); Assert.IsFalse(log.Contains("create")); }
            else { Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => run())); Assert.AreEqual(failure == 6 ? 1 : 0, log.FindAll(x => x == "create").Count); }
        }

        internal static IsolatedTestDesktop.MainInventory Inventory() => new IsolatedTestDesktop.MainInventory {
            Desktop = "Default", Complete = true, Visited = 4, Windows = new[] {
                new IsolatedTestDesktop.MainWindow { Handle = 11, ProcessId = 42, ThreadId = 7, ClassName = "OpusApp", Visible = true },
                new IsolatedTestDesktop.MainWindow { Handle = 12, ProcessId = 42, ThreadId = 7, ClassName = "wndclass_desked_gsk", Visible = true }
            } };

        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(6)]
        [DataRow(7)] [DataRow(8)] [DataRow(9)] [DataRow(10)] [DataRow(11)] [DataRow(12)] [DataRow(13)]
        [DataRow(14)] [DataRow(15)] [DataRow(16)] [DataRow(17)] [DataRow(18)]
        public void MainInventoryRejectsPartialForeignHiddenMissingOrAmbiguousEvidence(int fault)
        {
            var data = Inventory(); uint pid = 42; IntPtr root = new IntPtr(11);
            switch (fault)
            {
                case 1: data = null; break;
                case 2: data.Complete = false; break;
                case 3: data.Desktop = "Private"; break;
                case 4: data.Visited = -1; break;
                case 5: data.Visited = 8193; break;
                case 6: data.Visited = 1; break;
                case 7: data.Windows = null; break;
                case 8: data.Windows[0] = null; break;
                case 9: data.Windows[0].Handle = 0; break;
                case 10: data.Windows[0].ProcessId = 43; break;
                case 11: data.Windows[0].ThreadId = 0; break;
                case 12: data.Windows[0].ClassName = null; break;
                case 13: data.Windows[1].Handle = 11; break;
                case 14: data.Windows[0].Visible = false; break;
                case 15: data.Windows[1].Visible = false; break;
                case 16: root = new IntPtr(13); break;
                case 17: pid = 0; break;
                case 18: data.Windows = new[] { data.Windows[0], data.Windows[1],
                    new IsolatedTestDesktop.MainWindow { Handle = 13, ProcessId = 42, ThreadId = 7, ClassName = "wndclass_desked_gsk", Visible = true } }; break;
            }
            if (fault == 0) IsolatedTestDesktop.RequireMainInventory(data, pid, root, true, true);
            else Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireMainInventory(data, pid, root, true, true));
        }

        [TestMethod]
        public void EmptyCompleteInventoryIsOnlyAllowedWhenNoWindowIsRequired()
        {
            var empty = new IsolatedTestDesktop.MainInventory { Desktop = "Default", Complete = true, Windows = new IsolatedTestDesktop.MainWindow[0] };
            IsolatedTestDesktop.RequireMainInventory(empty, 42, IntPtr.Zero, false, false);
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireMainInventory(empty, 42, IntPtr.Zero, true, false));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireMainInventory(empty, 42, new IntPtr(1), false, false));
        }
    }
}
