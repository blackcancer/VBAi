namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass,TestCategory("Unit"),DoNotParallelize]
    public sealed class VbeFormHistoryWorkflowTests
    {
        [TestMethod]
        public void DesignerHistoryRequiresActionVersionAndAvailableNativeStack()
        {
            using(var f=new FormsWorkflowFixture())
            {
                var r=f.Request("wrong");Assert.ThrowsException<ArgumentException>(()=>f.Service.NativeHistory(r));
                r.Action="undo";r.ExpectedTreeVersion=null;Assert.ThrowsException<ArgumentException>(()=>f.Service.NativeHistory(r));
                r.ExpectedTreeVersion="stale";Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeHistory(r));
                r=f.Request("undo");f.Form.Designer.CanUndo=false;r.ExpectedTreeVersion=f.Request().ExpectedTreeVersion;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeHistory(r));
                f.Form.Designer.CanRedo=false;r=f.Request("redo");Assert.ThrowsException<InvalidOperationException>(()=>f.Service.NativeHistory(r));
            }
        }
        [TestMethod]
        public void DesignerHistoryDiffersFromAvailabilityAndPreservesErrorsAcrossReadback()
        {
            foreach(int scenario in new[]{0,1,2,3,4,5})
            using(var f=new FormsWorkflowFixture())
            {
                string action=scenario==0?"redo":"undo";
                Action execute=()=>{
                    if(scenario==2||scenario==3)f.Form.FailDesigner=true;
                    if(scenario==4)f.Form.Designer.Caption="changed";
                    if(scenario==1||scenario==3||scenario==4)throw new InvalidOperationException("history rejected");
                    if(scenario==0)f.Form.Designer.Caption="redone";
                    if(scenario==5)f.Form.Designer.CanUndo=false;
                };
                f.Form.Designer.OnUndo=execute;f.Form.Designer.OnRedo=execute;
                dynamic result=f.Service.NativeHistory(f.Request(action));
                Assert.AreEqual(scenario==0,(bool)result.Verified);
                Assert.AreEqual(scenario!=0&&scenario!=4,(bool)result.VerificationPending);
                Assert.AreEqual(scenario==0||scenario==5,(bool)result.Executed);
                if(scenario==2||scenario==3){Assert.IsNull(result.Tree);Assert.IsNull(result.ReadErrorsAfter);}
                else Assert.IsNotNull(result.Tree);
                if(scenario==1||scenario==3||scenario==4)Assert.AreEqual("history rejected",(string)result.NativeError);
                if(scenario==2)Assert.AreEqual("designer unavailable",(string)result.NativeError);
            }
        }
    }
}
