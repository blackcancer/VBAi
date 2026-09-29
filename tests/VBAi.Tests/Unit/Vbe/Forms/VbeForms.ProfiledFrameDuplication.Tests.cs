namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        [TestMethod]
        public void ProfiledFramePlanAndCopyCoverEveryPositiveChildProfile()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Label1").Caption = "Heading";
            f.Source.Controls.AddExisting("TextBox", "Text1").Value = "payload";
            f.Source.Controls.AddExisting("CheckBox", "Check1").Value = true;
            f.Source.Controls.AddExisting("CommandButton", "Button1").Caption = "Run";
            f.Source.Controls.AddExisting("ComboBox", "Combo1").ListWidth = "90 pt";
            f.Source.Controls.AddExisting("OptionButton", "Option1").Value = true;
            var request = f.Request(name: "FrameCopy");
            dynamic plan = f.Service.FrameProfileCopyPlan(request);
            Assert.IsTrue((bool)plan.EligibleForLimitedProbe);
            Assert.AreEqual(6, (int)plan.DirectChildCount);
            Assert.IsTrue((bool)plan.ReadOnly);
            dynamic result = f.Service.DuplicateFrameProfiled(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual(6, (int)result.DirectChildrenCopied);
            Assert.AreEqual(6, copy.Controls.Count);
            Assert.AreEqual("payload", copy.Controls.Item("FrameCopy_Text1").Value);
            Assert.AreEqual(true, copy.Controls.Item("FrameCopy_Check1").Value);
            Assert.AreEqual("90 pt", copy.Controls.Item("FrameCopy_Combo1").ListWidth);
            Assert.AreEqual(false, copy.Controls.Item("FrameCopy_Option1").Value);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void ProfiledFramePlanRefusesUnprofiledChildAndBadValue()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("ToggleButton", "Toggle1");
            dynamic unsupported = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)unsupported.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)unsupported.Issues).Cast<string>().Any(issue => issue.Contains("no positive copy profile")));
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            f.Source.Controls.Remove("Toggle1");
            f.Source.Controls.AddExisting("CheckBox", "Check1").Value = null;
            dynamic invalid = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)invalid.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)invalid.Issues).Cast<string>().Any(issue => issue.Contains("Boolean")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledFrameRollsBackChildFailureAndReportsIncompleteRollback()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.FailNextChildCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            f.Form.Designer.Controls.FailNextCaption = true;
            f.Form.Designer.Controls.FailNextRemove = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            StringAssert.Contains(error.Message, "rollback was incomplete");
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledPlanReportsEmptyFrameNameCollisionAndInvalidGeometry()
        {
            var f = Create("Frame");
            dynamic empty = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)empty.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)empty.Issues).Cast<string>().Any(issue => issue.Contains("duplicate_empty_form_frame")));
            f.Source.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.AddExisting("Label", "FrameCopy_Title");
            dynamic collision = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)collision.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>().Any(issue => issue.Contains("proposed child name already exists")));
            f.Source.Width = 0;
            dynamic badGeometry = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)badGeometry.Issues).Cast<string>().Any(issue => issue.Contains("geometry")));
        }

        [TestMethod]
        public void ProfiledPlanRefusesUnreadableChildProfilesWithoutMutation()
        {
            var f = Create("Frame");
            var label = f.Source.Controls.AddExisting("Label", "Title");
            label.Font.Size = 250;
            dynamic fontIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)fontIssue.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)fontIssue.Issues).Cast<string>().Any(issue => issue.Contains("font is outside")));
            label.Font.Size = 10;
            f.Source.Controls.AddExisting("TextBox", "Entry").Value = 42;
            dynamic textIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)textIssue.Issues).Cast<string>().Any(issue => issue.Contains("TextBox.Value must be text")));
            f.Source.Controls.Item("Entry").Value = "good";
            f.Source.Controls.AddExisting("ComboBox", "Choices").ListWidth = "";
            dynamic comboIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)comboIssue.Issues).Cast<string>().Any(issue => issue.Contains("ListWidth must be nonempty")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledPlanRejectsStaleTreeAndDirectChildLimit()
        {
            var f = Create("Frame");
            var request = f.Request(name: "FrameCopy");
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameProfileCopyPlan(request));
            for (int index = 0; index < 129; index++)
                f.Source.Controls.AddExisting("Label", "Label" + index);
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameProfileCopyPlan(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using System.Reflection;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod] public void ProfiledFrameChecksPreflightGeometryFontsAndConcurrentTreeChanges() { FramePreflightFailures("profiled"); }
        [TestMethod] public void ProfiledFrameChecksEverySupportedReadbackAndHierarchyField() { FrameReadbackFailures("profiled"); }
        [TestMethod] public void ProfiledFrameReportsNativeFailuresAndIncompleteChildOrRootRollback() { FrameNativeAndRollbackFailures("profiled"); }

        [TestMethod]
        public void ProfiledPlanValidatesRootPrefixVersionControlTypeNamesAndDirectOwnership()
        {
            var f=FrameWithChildren("profiled");
            foreach(var path in new[]{null,"Controls/Frame1/Controls/Title","Pages/Frame1","Controls/missing"}) {
                var r=f.Request();r.ControlPath=path;
                if(path=="Controls/missing") Assert.ThrowsException<InvalidOperationException>(()=>f.Service.FrameProfileCopyPlan(r));
                else Assert.ThrowsException<ArgumentException>(()=>f.Service.FrameProfileCopyPlan(r));
            }
            var missing=f.Request();missing.ExpectedTreeVersion=null;
            Assert.ThrowsException<ArgumentException>(()=>f.Service.FrameProfileCopyPlan(missing));
            f.Form.Designer.Controls.AddExisting("Label","Wrong");var wrong=f.Request();wrong.ControlPath="Controls/Wrong";
            Assert.ThrowsException<InvalidOperationException>(()=>f.Service.FrameProfileCopyPlan(wrong));
            dynamic collision=f.Service.FrameProfileCopyPlan(f.Request("Frame1"));
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>().Any(issue=>issue.StartsWith("New Frame name already exists",StringComparison.Ordinal)));
            dynamic longNames=f.Service.FrameProfileCopyPlan(f.Request(new string('F',41)));
            Assert.IsTrue(((IEnumerable)longNames.Issues).Cast<string>().Any(issue=>issue=="New Frame name exceeds 40 characters."));
            Assert.IsTrue(((IEnumerable)longNames.Issues).Cast<string>().Any(issue=>issue.Contains("proposed child name exceeds")));
            f.Frame.Controls.Item("Title").Parent=f.Form.Designer;
            dynamic indirect=f.Service.FrameProfileCopyPlan(f.Request());
            Assert.AreEqual(0,(int)indirect.DirectChildCount);
            Assert.IsFalse((bool)indirect.EligibleForLimitedProbe);
        }

        [TestMethod]
        public void ProfiledPlanCoversNullableTextFalseCheckAndInvalidComboWidths()
        {
            foreach(var type in new[]{"TextBox","CheckBox"}) {
                var f=FrameWithChildren("profiled",type);
                dynamic result=f.Service.DuplicateFrameProfiled(f.Request());
                Assert.AreEqual(1,(int)result.DirectChildrenCopied);
                Assert.AreEqual(type=="TextBox"?null:(object)false,f.Form.Designer.Controls.Item("FrameCopy").Controls.Item("FrameCopy_Title").Value);
            }
            foreach(var width in new object[]{null,42," ",new string('x',65)}) {
                var f=FrameWithChildren("profiled","ComboBox");f.Frame.Controls.Item("Title").ListWidth=width;
                dynamic plan=f.Service.FrameProfileCopyPlan(f.Request());
                Assert.IsFalse((bool)plan.EligibleForLimitedProbe);
                Assert.IsTrue(((IEnumerable)plan.Issues).Cast<string>().Any(issue=>issue.Contains("ListWidth must be nonempty text")));
            }
        }

        [TestMethod]
        public void ProfiledFrameRejectsReadbackForEachNonLabelProfile()
        {
            foreach(var type in new[]{"TextBox","CheckBox","CommandButton","OptionButton","ComboBox"}) {
                foreach(var field in type=="CheckBox"?new[]{"Caption","ValueType","Value"}:new[]{type=="TextBox"?"Value":type=="ComboBox"?"ListWidth":"Caption"}) {
                    var f=FrameWithChildren("profiled",type);var request=f.Request();
                    f.Form.Designer.Controls.ConfigureAdded=frame=>frame.Controls.ConfigureAdded=c=>{
                        if(field=="ValueType") c.ReadOverrides["Value"]="invalid Boolean";
                        else c.ReadOverrides[field]=field=="Value"&&type=="CheckBox"?(object)true:"changed";
                    };
                    AssertRolledBack(f,request,"profiled","Copied "+(type=="CommandButton"||type=="OptionButton"?"button caption":type=="TextBox"?"TextBox value":type=="ComboBox"?"ComboBox ListWidth":type)+" differs");
                }
            }
        }

        [TestMethod]
        public void ProfileHelpersRejectUnregisteredTypesForReadApplyAndVerify()
        {
            var production=typeof(VBAi.VbeForms);
            var profileType=production.GetNestedType("VerifiedChildProfile",BindingFlags.NonPublic);
            var snapshotType=production.GetNestedType("ProfiledChildSnapshot",BindingFlags.NonPublic);
            var profile=Activator.CreateInstance(profileType,true);
            profileType.GetField("Type").SetValue(profile,"Unknown");
            var snapshot=Activator.CreateInstance(snapshotType,true);
            snapshotType.GetField("Profile").SetValue(snapshot,profile);
            snapshotType.GetField("Width").SetValue(snapshot,20d);
            snapshotType.GetField("Height").SetValue(snapshot,10d);
            var f=Create();var child=f.Frame.Controls.AddExisting("Label","Title");
            var read=production.GetMethod("ReadProfiledChild",BindingFlags.NonPublic|BindingFlags.Static);
            var apply=production.GetMethod("ApplyProfiledChild",BindingFlags.NonPublic|BindingFlags.Static);
            var verify=production.GetMethod("VerifyProfiledChild",BindingFlags.NonPublic|BindingFlags.Static);
            foreach(var invocation in new Action[]{
                ()=>read.Invoke(null,new object[]{child,profile,"source","proposed","name"}),
                ()=>apply.Invoke(null,new object[]{child,snapshot}),
                ()=>verify.Invoke(null,new object[]{child,snapshot})}) {
                var error=Assert.ThrowsException<TargetInvocationException>(invocation);
                Assert.IsInstanceOfType(error.InnerException,typeof(InvalidOperationException));
                StringAssert.Contains(error.InnerException.Message,"Profile has no");
            }
        }
    }
}
