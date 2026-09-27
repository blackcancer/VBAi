namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void FrameCopyPlanDescribesDirectLabelsWithoutMutatingTheForm()
        {
            var f = Create();
            var label = f.Frame.Controls.AddExisting("Label", "Title");
            label.Caption = "Hello";
            var request = f.Request();
            dynamic plan = f.Service.FrameCopyPlan(request);
            Assert.IsTrue((bool)plan.EligibleForLimitedProbe);
            Assert.IsTrue((bool)plan.ReadOnly);
            Assert.IsFalse((bool)plan.MutationVerified);
            Assert.AreEqual(1, (int)plan.CollectionCount);
            Assert.AreEqual(1, (int)plan.DirectChildCount);
            var child = ((IEnumerable)plan.Children).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1/Controls/Title", (string)((dynamic)child).SourcePath);
            Assert.AreEqual("Controls/FrameCopy/Controls/FrameCopy_Title", (string)((dynamic)child).ProposedPath);
            Assert.AreEqual("Label positive profile", (string)((dynamic)child).Profile);
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual((string)request.ExpectedTreeVersion, (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void SimplePlanAcceptsTextBoxButLabelOnlyPlanRefusesIt()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("TextBox", "Entry");
            var request = f.Request();
            dynamic labels = f.Service.FrameCopyPlan(request);
            dynamic simple = f.Service.FrameSimpleCopyPlan(request);
            Assert.IsFalse((bool)labels.EligibleForLimitedProbe);
            Assert.IsFalse((bool)((dynamic)((IEnumerable)labels.Children).Cast<object>().Single()).Eligible);
            StringAssert.Contains((string)((dynamic)((IEnumerable)labels.Children).Cast<object>().Single()).Issue, "Only direct Label");
            Assert.IsTrue((bool)simple.EligibleForLimitedProbe);
            Assert.AreEqual("TextBox text-value positive profile", (string)((dynamic)((IEnumerable)simple.Children).Cast<object>().Single()).Profile);
        }

        [TestMethod]
        public void PlanReportsNameCollisionAndOversizedChildNameWithoutWriting()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.AddExisting("Label", "FrameCopy_Title");
            var request = f.Request();
            dynamic collision = f.Service.FrameCopyPlan(request);
            Assert.IsFalse((bool)collision.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>().Any(issue => issue.Contains("already exists")));
            request.NewName = new string ('F', 38);
            dynamic longChild = f.Service.FrameSimpleCopyPlan(request);
            Assert.IsFalse((bool)longChild.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)longChild.Issues).Cast<string>().Any(issue => issue.Contains("40 characters")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void PlanRejectsStalePathWrongTypeAndInvalidName()
        {
            var f = Create();
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            request.ControlPath = "Controls/frame1";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            request.ControlPath = "Controls/Frame1";
            request.NewName = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => f.Service.FrameCopyPlan(request));
            request.NewName = "FrameCopy";
            var label = f.Form.Designer.Controls.AddExisting("Label", "Label1");
            request.ControlPath = "Controls/Label1";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            Assert.AreEqual("Label1", label.Name);
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void FramePlansValidateRequiredPathsAndReportBothNameIssues()
        {
            var f=Create();
            foreach(var field in new[]{"ControlPath","ExpectedTreeVersion"}) {
                var r=f.Request();typeof(CodexVBE.Request).GetProperty(field).SetValue(r,null);
                Assert.ThrowsException<ArgumentException>(()=>f.Service.FrameCopyPlan(r));
            }
            dynamic collision=f.Service.FrameCopyPlan(f.Request("Frame1"));
            Assert.IsFalse((bool)collision.EligibleForLimitedProbe);
            StringAssert.Contains(((IEnumerable)collision.Issues).Cast<string>().Single(),"New Frame name already exists");
            dynamic longName=f.Service.FrameSimpleCopyPlan(f.Request(new string('F',41)));
            Assert.IsFalse((bool)longName.EligibleForLimitedProbe);
            StringAssert.Contains(((IEnumerable)longName.Issues).Cast<string>().Single(),"exceeds 40 characters");
            f.Frame.Controls.AddExisting("CheckBox","Unsupported");
            dynamic simple=f.Service.FrameSimpleCopyPlan(f.Request());
            StringAssert.Contains(((IEnumerable)simple.Issues).Cast<string>().Single(),"Label or TextBox");
        }

        [TestMethod]
        public void FramePlanSkipsNonDirectControlsAndBoundsChangesAfterTreeRead()
        {
            var f=Create();
            var child=f.Frame.Controls.AddExisting("Label","OtherParent");
            child.Parent=f.Form.Designer;
            dynamic plan=f.Service.FrameCopyPlan(f.Request());
            Assert.AreEqual(0,(int)plan.DirectChildCount);
            Assert.AreEqual(1,(int)plan.CollectionCount);
            child.Parent=f.Frame;
            var r=f.Request();int reads=0;
            f.Frame.Controls.BeforeEnumeration=()=>{
                if(++reads==2) for(int i=0;i<512;i++) f.Frame.Controls.AddExisting("Label","L"+i);
            };
            var error=Assert.ThrowsException<InvalidOperationException>(()=>f.Service.FrameCopyPlan(r));
            StringAssert.Contains(error.Message,"more than 512 direct controls");
            Assert.AreEqual(1,f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void FramePlanRejectsCanonicalPageRatherThanControl()
        {
            var f=Create(); var page=f.Frame.Pages.Add("Page1","Page");
            var r=f.Request();r.ControlPath="Controls/Frame1/Pages/Page1";
            Assert.ThrowsException<ArgumentException>(()=>f.Service.FrameCopyPlan(r));
        }
    }
}
