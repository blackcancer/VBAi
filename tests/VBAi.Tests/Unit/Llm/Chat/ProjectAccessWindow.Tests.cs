using System.Collections.Generic;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ProjectAccessWindowTests
    {
        [STATestMethod]
        public void StaticDesignerControlsRetainSelectionsAndSeparateSharedPermission()
        {
            using (var window = new ProjectAccessWindow())
            {
                window.Populate(new[] {
                    new KeyValuePair<string, string>("C:\\B.xlsm", "B"),
                    new KeyValuePair<string, string>("C:\\C.xlsm", "C") }, new[] { "C:\\B.xlsm" }, false);
                CollectionAssert.AreEqual(new[] { "C:\\B.xlsm" }, window.SelectedProjects);
                Assert.IsFalse(window.SharedContext);
                var list = (CheckedListBox)window.Controls["projectList"];
                list.SetItemChecked(1, true);
                Assert.AreEqual(2, window.SelectedProjects.Length);
                ((CheckBox)window.Controls["sharedContext"]).Checked = true;
                Assert.IsTrue(window.SharedContext);
                Assert.AreEqual(DialogResult.OK, ((Button)window.AcceptButton).DialogResult);
                Assert.AreEqual(DialogResult.Cancel, ((Button)window.CancelButton).DialogResult);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using System.Threading;
    using System.Reflection;
    using System.Windows.Forms;
    [TestClass,TestCategory("Unit")]
    public sealed class ProjectAccessAppearanceTests
    {
        [STATestMethod]
        public void ProjectDialogHandlesDesignerNullGrantsOwnedThreadDispatchAndDisposal()
        {
            using(var designed=(ProjectAccessWindow)LicenseManager.CreateWithContext(typeof(ProjectAccessWindow),new AccessDesignContext())) {
                designed.Populate(new[]{new KeyValuePair<string,string>("A","Alpha")},null,true);
                Assert.AreEqual("Alpha",((CheckedListBox)designed.Controls["projectList"]).Items[0].ToString());
                Assert.AreEqual(0,designed.SelectedProjects.Length);Assert.IsTrue(designed.SharedContext);
            }
            var window=new ProjectAccessWindow();
            try {
                window.Populate(new[]{new KeyValuePair<string,string>("B","Beta")},new[]{"b"},false);
                CollectionAssert.AreEqual(new[]{"B"},window.SelectedProjects);
                var owned=window.Handle;
                var worker=new Thread(()=>AccessContract.Call(window,"ApplyAppearance"));worker.Start();Assert.IsTrue(worker.Join(3000));
                Application.DoEvents();
                AccessContract.Call(window,"ApplyAppearance");
                typeof(ProjectAccessWindow).GetMethod("Dispose",BindingFlags.Instance|BindingFlags.NonPublic,null,new[]{typeof(bool)},null).Invoke(window,new object[]{false});
            } finally { window.Dispose(); }
            AccessContract.Call(window,"ApplyAppearance");
            var noContainer=new ProjectAccessWindow();
            var field=typeof(ProjectAccessWindow).GetField("components",BindingFlags.Instance|BindingFlags.NonPublic);
            var container=(IContainer)field.GetValue(noContainer);field.SetValue(noContainer,null);
            noContainer.Dispose();container.Dispose();
        }
    }
}
