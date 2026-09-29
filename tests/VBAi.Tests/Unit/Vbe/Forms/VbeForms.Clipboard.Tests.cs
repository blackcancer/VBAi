using System;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.Linq;
namespace VBAi.Tests.Unit
{
    [TestClass,TestCategory("Unit"),DoNotParallelize]
    public sealed class VbeFormClipboardWorkflowTests
    {
        [TestMethod]
        public void ClipboardStateReadsRootAndNestedSelectionsAndRefusesChangingClipboardOrExcessSelection()
        {
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;
                dynamic state=f.Service.ClipboardState(f.Project.Name,f.Form.Name);
                CollectionAssert.AreEqual(new[]{"A"},((IEnumerable)state.Selected).Cast<string>().ToArray());
                Assert.AreEqual("10",(string)state.ClipboardVersion);
                var frame=f.Form.Designer.Controls.Add("Frame","Frame");frame.Controls.Add("Label","Child").InSelection=true;
                dynamic nested=f.Service.ClipboardState(f.Project.Name,f.Form.Name,"Controls/Frame");
                CollectionAssert.AreEqual(new[]{"Child"},((IEnumerable)nested.Selected).Cast<string>().ToArray());
                Assert.ThrowsException<ArgumentException>(()=>f.Service.ClipboardState(f.Project.Name,f.Form.Name,"Controls/Missing"));
                VbeForms.DesignerClipboardSequence=()=>++f.Sequence;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.ClipboardState(f.Project.Name,f.Form.Name));
            }
            using(var f=new FormsWorkflowFixture())
            {
                foreach(var node in f.Form.Designer.Controls.Items)node.InSelection=true;
                for(int i=0;i<255;i++)f.Form.Designer.Controls.Add("Label","C"+i).InSelection=true;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.ClipboardState(f.Project.Name,f.Form.Name));
            }
        }
        [TestMethod]
        public void ClipboardNativeActionsValidateRevisionSelectionAndPasteAvailability()
        {
            using(var f=new FormsWorkflowFixture())
            {
                Assert.ThrowsException<ArgumentException>(()=>f.Service.NativeClipboard(f.Request("wrong")));
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeClipboard(f.ClipboardRequest("copy")));
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeClipboard(f.ClipboardRequest("cut")));
                f.Form.Designer.CanPaste=false;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeClipboard(f.ClipboardRequest("paste")));
                var request=f.ClipboardRequest();request.ExpectedClipboardVersion=null;
                Assert.ThrowsException<InvalidOperationException>(()=>VbeForms.RequireClipboardRevision(request,request.ExpectedDesignerSelectionVersion,"10"));
            }
        }
        [TestMethod]
        public void ClipboardCopyAndPasteReportSeparateDesignerAndClipboardChanges()
        {
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;
                dynamic copied=f.Service.NativeClipboard(f.ClipboardRequest("copy"));
                Assert.IsTrue((bool)copied.Executed);Assert.IsTrue((bool)copied.ClipboardChanged);Assert.IsFalse((bool)copied.DesignerChangeObserved);
                Assert.IsNull(copied.DesignerClipboardRecoveryId);Assert.IsFalse((bool)copied.ClipboardContentVerified);
                dynamic pasted=f.Service.NativeClipboard(f.ClipboardRequest("paste"));
                Assert.IsTrue((bool)pasted.Executed);Assert.IsTrue((bool)pasted.DesignerChangeObserved);Assert.IsFalse((bool)pasted.ClipboardChanged);
                Assert.AreEqual(3,f.Form.Designer.Controls.Count);
            }
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;f.Form.Designer.OnCopy=()=>{};
                Assert.IsFalse((bool)((dynamic)f.Service.NativeClipboard(f.ClipboardRequest("copy"))).ClipboardChanged);
            }
        }
        [TestMethod]
        public void ClipboardActionsPreserveExecutionErrorAndHandleReadbackFailure()
        {
            foreach(int scenario in new[]{0,1,2})
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;
                f.Form.Designer.OnCopy=()=>{if(scenario!=0)f.Project.Mode=0;if(scenario!=1)throw new InvalidOperationException("copy rejected");};
                dynamic result=f.Service.NativeClipboard(f.ClipboardRequest("copy"));
                Assert.IsFalse((bool)result.Executed);Assert.AreEqual(scenario==1?"The project must be in design mode.":"copy rejected",(string)result.NativeError);
                if(scenario!=0){Assert.IsNull(result.After);Assert.IsNull(result.ReadErrorsAfter);}else Assert.IsNotNull(result.After);
            }
        }
        [TestMethod]
        public void CutCapturesBeforeRemovalAndRejectsUnstableDataAndPartialFailures()
        {
            foreach(int scenario in new[]{0,1,2,3,4})
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;
                if(scenario==0)f.OnReadClipboard=()=>f.ClipboardData=null;
                if(scenario==1)f.OnReadClipboard=()=>f.Sequence++;
                if(scenario==2)f.Form.Designer.OnCut=()=>{throw new InvalidOperationException("cut rejected");};
                if(scenario==3)f.Form.Designer.OnCut=()=>{};
                if(scenario==4)f.Form.Designer.OnCut=()=>{f.Cut(f.Form.Designer);f.Project.Mode=0;};
                dynamic result=f.Service.NativeClipboard(f.ClipboardRequest("cut"));
                if(scenario<2)Assert.IsNull(result.DesignerClipboardRecoveryId);else Assert.IsNotNull(result.DesignerClipboardRecoveryId);
                Assert.AreEqual(scenario==3,(bool)result.Executed);
                if(scenario==4)Assert.IsNull(result.After);
            }
        }
        [TestMethod]
        public void CutRecoveriesRetainLatestEightAndRestoreOnlyTheCapturedClipboard()
        {
            using(var f=new FormsWorkflowFixture())
            {
                string first=null,last=null;
                for(int i=0;i<9;i++)
                {
                    int index=0;foreach(var node in f.Form.Designer.Controls.Items){node.InSelection=false;node.tab=index++;}
                    if(i==0)f.Form.Designer.Controls.Item("A").InSelection=true;
                    else f.Form.Designer.Controls.Add("Label","A"+i).InSelection=true;
                    dynamic cut=f.Service.NativeClipboard(f.ClipboardRequest("cut"));last=cut.DesignerClipboardRecoveryId;if(i==0)first=last;
                    Assert.IsTrue((bool)cut.Executed);Assert.IsTrue((bool)cut.DesignerChangeObserved);Assert.IsNotNull(last);
                }
                var r=f.ClipboardRequest();r.DesignerClipboardRecoveryId=first;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RestoreDesignerClipboard(r));
                r.DesignerClipboardRecoveryId=last;
                dynamic restored=f.Service.RestoreDesignerClipboard(r);Assert.IsTrue((bool)restored.Restored);Assert.AreEqual(1,f.Form.Designer.Controls.Count);
                r=f.ClipboardRequest();r.DesignerClipboardRecoveryId=last;
                f.OnReadClipboard=()=>f.ClipboardData=null;
                Assert.IsFalse((bool)((dynamic)f.Service.RestoreDesignerClipboard(r)).Restored);
            }
        }
        [TestMethod]
        public void ClipboardRestoreRefusesOtherLiveFormAndContainer()
        {
            using(var f=new FormsWorkflowFixture())
            {
                f.Form.Designer.Controls.Item("A").InSelection=true;dynamic cut=f.Service.NativeClipboard(f.ClipboardRequest("cut"));
                var r=f.ClipboardRequest();r.DesignerClipboardRecoveryId=cut.DesignerClipboardRecoveryId;r.ParentPath="Controls/Frame";
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RestoreDesignerClipboard(r));
                r.ParentPath=null;f.Project.VBComponents[0]=new FormsWorkflowFixture.WorkflowForm(f);
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RestoreDesignerClipboard(r));
            }
        }
        [TestMethod]
        public void DesignerSelectionValidatesNamesVersionAndContainerAndCanClearSelection()
        {
            using(var f=new FormsWorkflowFixture())
            {
                foreach(var items in new[]{(string[])null,new string[65],new[]{" "},new[]{"A","A"}})
                {var r=f.ClipboardRequest();r.Items=items;Assert.ThrowsException<ArgumentException>(()=>f.Service.SelectDesignerControls(r));}
                foreach(string version in new[]{(string)null,"stale"})
                {var r=f.ClipboardRequest();r.Items=new[]{"A"};r.ExpectedDesignerSelectionVersion=version;Assert.ThrowsException<InvalidOperationException>(()=>f.Service.SelectDesignerControls(r));}
                var request=f.ClipboardRequest();request.Items=new[]{"missing"};Assert.ThrowsException<ArgumentException>(()=>f.Service.SelectDesignerControls(request));
                request=f.ClipboardRequest();request.Items=new[]{"A"};Assert.IsTrue((bool)((dynamic)f.Service.SelectDesignerControls(request)).Verified);
                Assert.IsTrue(f.Form.Designer.Controls.Item("A").InSelection);Assert.IsFalse(f.Form.Designer.Controls.Item("B").InSelection);
                request=f.ClipboardRequest();request.Items=new string[0];Assert.IsTrue((bool)((dynamic)f.Service.SelectDesignerControls(request)).Verified);
                Assert.IsFalse(f.Form.Designer.Controls.Item("A").InSelection);
            }
        }
        [TestMethod]
        public void DesignerSelectionNeverClaimsSuccessOnIgnoredSettersOrPartialReadbackFailure()
        {
            foreach(int scenario in new[]{0,1,2,3})
            using(var f=new FormsWorkflowFixture())
            {
                var a=f.Form.Designer.Controls.Item("A");
                a.OnSelection=value=>{if(scenario>=2)f.Project.Mode=0;if(scenario==1||scenario==3)throw new InvalidOperationException("selection rejected");};
                if(scenario==0)a.IgnoreSelection=true;
                var r=f.ClipboardRequest();r.Items=new[]{"A"};dynamic result=f.Service.SelectDesignerControls(r);
                Assert.IsFalse((bool)result.Verified);
                if(scenario>=2)Assert.IsNull(result.After);
                if(scenario==1||scenario==3)Assert.AreEqual("selection rejected",(string)result.NativeError);
            }
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed class VbeFormClipboardTests
    {
        [TestMethod]
        public void ClipboardAndSelectionRevisionsAreBothRequired()
        {
            var request = new Request { ExpectedDesignerSelectionVersion = "tree-selection", ExpectedClipboardVersion = "123" };
            VbeForms.RequireClipboardRevision(request, "tree-selection", "123");
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "changed", "123"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "124"));
            request.ExpectedClipboardVersion = "0";
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "0"));
            request.ExpectedDesignerSelectionVersion = null;
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "123"));
        }
    }
}
