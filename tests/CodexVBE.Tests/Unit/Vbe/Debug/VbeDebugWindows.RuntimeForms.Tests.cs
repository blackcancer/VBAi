namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeRuntimeFormsTests
    {
        [TestMethod]
        public void RuntimeObservationFiltersOtherProcessesHiddenAndUnrelatedClasses()
        {
            using (var scene = new NativeDebugScene())
            {
                scene.Add("ThunderDFrame", "foreign").Process = uint.MaxValue;
                scene.Add("ThunderDFrame", "hidden").Visible = false;
                var owner = scene.Add("other", "owner"); owner.Enabled = false;
                var form = scene.Add("ThunderDFrame", "visible", left: 5, top: 6, width: 70, height: 80);
                form.Owner = owner.Handle;
                var noBounds = scene.Add("ThunderDFrame", "no bounds"); noBounds.BoundsAvailable = false;
                scene.Add("ThunderDFrame", "broken").FailBounds = true;
                scene.EnumerationSucceeded = false;
                dynamic result = VbeDebugWindows.ReadRuntimeForms();
                object[] rows = ((IEnumerable)result.Windows).Cast<object>().ToArray();
                dynamic row = rows[0]; dynamic bounds = row.Bounds;
                Assert.AreEqual(2, rows.Length); Assert.AreEqual(3, (int)result.DetectedCount);
                Assert.AreEqual("visible", (string)row.Caption); Assert.AreEqual(owner.Handle.ToInt64(), (long)row.OwnerHandle);
                Assert.AreEqual(false, (bool)row.OwnerEnabled); Assert.AreEqual(70, (int)bounds.Width);
                Assert.AreEqual(80, (int)bounds.Height); Assert.AreEqual(5, (int)bounds.Left); Assert.AreEqual(6, (int)bounds.Top);
                dynamic absent = rows[1]; Assert.IsNull(absent.OwnerEnabled); Assert.IsNull(absent.Bounds);
                Assert.IsFalse((bool)result.EnumerationSucceeded); Assert.IsTrue((bool)result.Truncated);
                CollectionAssert.AreEqual(new[] { "bounds unavailable" }, ((IEnumerable)result.Errors).Cast<string>().ToArray());
                Assert.IsFalse((bool)row.ProjectIdentityVerified); Assert.IsNull(row.Project);
            }
        }
        [TestMethod]
        public void RuntimeObservationCapsRowsButCountsEveryDetectedWindow()
        {
            using (var scene = new NativeDebugScene())
            {
                dynamic empty = VbeDebugWindows.ReadRuntimeForms();
                Assert.IsFalse((bool)empty.Truncated); Assert.IsTrue((bool)empty.EnumerationSucceeded);
                for (int i = 0; i < 66; i++) scene.Add("ThunderDFrame", "form" + i);
                dynamic result = VbeDebugWindows.ReadRuntimeForms();
                Assert.AreEqual(66, (int)result.DetectedCount);
                Assert.AreEqual(64, ((IEnumerable)result.Windows).Cast<object>().Count());
                Assert.IsTrue((bool)result.Truncated);
            }
        }
    }
}
