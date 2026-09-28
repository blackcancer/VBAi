namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass,TestCategory("Unit"),DoNotParallelize]
    public sealed class VbeCutRecoveryWorkflowTests
    {
        private static string Cut(FormsWorkflowFixture f)
        {
            f.Form.Designer.Controls.Item("A").InSelection=true;
            dynamic cut=f.Service.NativeClipboard(f.ClipboardRequest("cut"));
            Assert.IsTrue((bool)cut.Executed,(string)cut.NativeError);return cut.DesignerClipboardRecoveryId;
        }
        private static Request Recovery(FormsWorkflowFixture f,string id)
        {var r=f.ClipboardRequest();r.DesignerClipboardRecoveryId=id;return r;}
        [TestMethod]
        public void RecoveryAvailabilityRequiresCompletedCutLiveIdentityContainerAndUnusedAttempt()
        {
            using(var f=new FormsWorkflowFixture())
            {
                var unknown=f.Request();unknown.DesignerClipboardRecoveryId="unknown";
                Assert.IsFalse(f.Service.CanRecoverCut(unknown));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(unknown));
                f.Form.Designer.OnCut=()=>{};string incomplete=Cut(f);var r=Recovery(f,incomplete);
                Assert.IsFalse(f.Service.CanRecoverCut(r));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(r));
                f.Form.Designer.OnCut=null;string id=Cut(f);r=Recovery(f,id);Assert.IsTrue(f.Service.CanRecoverCut(r));
                r.ParentPath="Controls/Other";Assert.IsFalse(f.Service.CanRecoverCut(r));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(r));
                r.ParentPath=null;r.Project="missing";Assert.IsFalse(f.Service.CanRecoverCut(r));
                r.Project=f.Project.Name;f.Project.VBComponents[0]=new FormsWorkflowFixture.WorkflowForm(f);
                Assert.IsFalse(f.Service.CanRecoverCut(r));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(r));
            }
        }
        [TestMethod]
        public void RecoveryRestoresNamesOriginalGeometryAndNativeTabOrderOnlyOnce()
        {
            using(var f=new FormsWorkflowFixture())
            {
                string id=Cut(f);var r=Recovery(f,id);
                dynamic result=f.Service.RecoverDesignerCut(r);
                Assert.IsTrue((bool)result.RestoredNamesGeometryAndTabOrder,(string)result.NativeError);
                Assert.IsTrue((bool)result.RecoveryAttempted);Assert.IsFalse((bool)result.FullPropertyFidelityVerified);
                Assert.AreEqual(0,f.Form.Designer.Controls.Item("A").Left);Assert.AreEqual(0,f.Form.Designer.Controls.Item("A").Top);
                Assert.AreEqual(0,f.Form.Designer.Controls.Item("A").TabIndex);Assert.AreEqual(1,f.Form.Designer.Controls.Item("B").TabIndex);
                Assert.IsFalse(f.Service.CanRecoverCut(r));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(r));
                Assert.AreEqual(2,f.Form.Designer.Controls.Count);
            }
        }
        [TestMethod]
        public void RecoveryRefusesInterveningEditsAndUnexpectedPartialNativeCutResults()
        {
            foreach(int scenario in new[]{0,1,2})
            using(var f=new FormsWorkflowFixture())
            {
                if(scenario==1)
                {
                    f.Form.Designer.Controls.Item("B").InSelection=true;
                    f.Form.Designer.OnCut=()=>f.Form.Designer.Controls.Items.RemoveAll(x=>x.Name=="A");
                }
                if(scenario==2)f.Form.Designer.OnCut=()=>{f.Cut(f.Form.Designer);f.Form.Designer.Controls.Items.Clear();};
                string id=Cut(f);
                if(scenario==0)f.Form.Designer.Caption="external edit";
                var error=Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(Recovery(f,id)));
                StringAssert.Contains(error.Message,scenario==0?"changed after the cut":"already present or the container changed");
            }
        }
        [TestMethod]
        public void RecoveryPreservesClipboardChecksAndDoesNotConsumeAnAttemptBeforePaste()
        {
            foreach(int scenario in new[]{0,1,2})
            using(var f=new FormsWorkflowFixture())
            {
                string id=Cut(f);var r=Recovery(f,id);
                if(scenario==0)f.OnWriteClipboard=()=>{throw new InvalidOperationException("publish rejected");};
                if(scenario==1)f.OnReadClipboard=()=>f.ClipboardData=null;
                if(scenario==2)f.Form.Designer.CanPaste=false;
                dynamic result=f.Service.RecoverDesignerCut(r);
                Assert.IsFalse((bool)result.RestoredNamesGeometryAndTabOrder);Assert.IsFalse((bool)result.RecoveryAttempted);
                StringAssert.Contains((string)result.NativeError,scenario==0?"publish rejected":scenario==1?"readback differs":"cannot paste");
                Assert.IsTrue(f.Service.CanRecoverCut(r));
            }
        }
        [TestMethod]
        public void RecoveryConsumesPasteAndReportsNamesGeometryAndTabReadbackFailures()
        {
            foreach(int scenario in new[]{0,1,2,3})
            using(var f=new FormsWorkflowFixture())
            {
                string id=Cut(f);var r=Recovery(f,id);
                f.Form.Designer.OnPaste=()=>{
                    if(scenario==0)throw new InvalidOperationException("paste rejected");
                    f.Paste(f.Form.Designer);var a=f.Form.Designer.Controls.Item("A");
                    if(scenario==1)a.Name="NativeRenamed";
                    if(scenario==2)a.OnSetGeometry=(name,value)=>{throw new InvalidOperationException("geometry rejected");};
                    if(scenario==3)a.IgnoreTab=true;
                };
                dynamic result=f.Service.RecoverDesignerCut(r);
                Assert.IsFalse((bool)result.RestoredNamesGeometryAndTabOrder);Assert.IsTrue((bool)result.RecoveryAttempted);
                StringAssert.Contains((string)result.NativeError,scenario==0?"paste rejected":scenario==1?"names differ":scenario==2?"geometry rejected":"tab order differs");
                Assert.IsFalse(f.Service.CanRecoverCut(r));Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RecoverDesignerCut(r));
            }
        }
        [TestMethod]
        public void RecoveryDoesNotClaimRestorationWhenFinalInspectionFailsAndPreservesPasteError()
        {
            foreach(bool pasteFails in new[]{false,true})
            using(var f=new FormsWorkflowFixture())
            {
                string id=Cut(f);var r=Recovery(f,id);
                f.Form.Designer.OnPaste=()=>{
                    if(pasteFails){f.Project.Mode=0;throw new InvalidOperationException("paste rejected");}
                    f.Paste(f.Form.Designer);
                    f.Form.Designer.Controls.Item("B").OnReadTab=value=>{f.Project.Mode=0;return value;};
                };
                dynamic result=f.Service.RecoverDesignerCut(r);
                Assert.IsFalse((bool)result.RestoredNamesGeometryAndTabOrder);Assert.IsTrue((bool)result.RecoveryAttempted);
                Assert.IsNull(result.State);Assert.IsNull(result.RemainingDifferences);Assert.IsNull(result.ReadErrorsAfter);
                Assert.AreEqual(pasteFails?"paste rejected":"The project must be in design mode.",(string)result.NativeError);
            }
        }
    }
}
