namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Win32;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed class VbeControlCatalogTests
    {
        [TestMethod]
        public void DisposableRegistryCatalogHandlesAbsentClassesAndIncompleteControlRegistrations()
        {
            string ownedKey = "Software\\CodexVBE\\Tests\\ControlCatalog\\" + Guid.NewGuid().ToString("N");
            var previous = VbeControlCatalog.OpenClasses;
            try
            {
                using (Registry.CurrentUser.CreateSubKey(ownedKey)) { }
                VbeControlCatalog.OpenClasses = () => Registry.CurrentUser.OpenSubKey(ownedKey);
                var native = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Forms.Label.1" };
                Assert.AreEqual(1, ((IEnumerable<object>)VbeControlCatalog.List(native)).Count());
                const string category = "\\Implemented Categories\\{40FC6ED4-2438-11CF-A3DB-080036F12502}";
                using (var root = Registry.CurrentUser.OpenSubKey(ownedKey, true))
                {
                    using (root.CreateSubKey("CLSID\\no-category\\ProgID")) { }
                    foreach (string name in new[] { "native", "installed", "no-progid", "empty", "wrong-type" }) using (root.CreateSubKey("CLSID\\" + name + category)) { }
                    using (var key = root.CreateSubKey("CLSID\\native\\ProgID")) key.SetValue("", "Forms.Label.1");
                    using (var key = root.CreateSubKey("CLSID\\installed\\ProgID")) key.SetValue("", "Coverage.DisposableControl.1");
                    using (var key = root.CreateSubKey("CLSID\\empty\\ProgID")) key.SetValue("", " ");
                    using (var key = root.CreateSubKey("CLSID\\wrong-type\\ProgID")) key.SetValue("", 5);
                }
                Assert.AreEqual(2, ((IEnumerable<object>)VbeControlCatalog.List(native)).Count());
                Assert.IsTrue(VbeControlCatalog.IsCandidate("Forms.Label.1", native));
                Assert.IsTrue(VbeControlCatalog.IsCandidate("coverage.disposablecontrol.1", native));
                Assert.IsFalse(VbeControlCatalog.IsCandidate("Coverage.Missing.1", native));
            }
            finally { VbeControlCatalog.OpenClasses = previous; Registry.CurrentUser.DeleteSubKeyTree(ownedKey); }
        }
    }
}
