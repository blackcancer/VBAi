using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        private static object CopyFrame(Fixture f, Request r, string route)
        {
            switch(route) {
                case "empty": return f.Service.DuplicateEmptyFrame(r);
                case "labels": return f.Service.DuplicateFrameWithLabels(r);
                case "simple": return f.Service.DuplicateFrameWithSimpleChildren(r);
                default: return f.Service.DuplicateFrameProfiled(r);
            }
        }
        private static Fixture FrameWithChildren(string route, string childType = "Label")
        {
            var f=Create();
            if(route!="empty") f.Frame.Controls.AddExisting(childType,"Title");
            if(route=="simple") f.Frame.Controls.AddExisting("TextBox","Entry").Value="payload";
            return f;
        }
        private static void AssertRolledBack(Fixture f, Request r, string route, string message)
        {
            var error=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(f,r,route));
            StringAssert.Contains(error.Message,message);
            Assert.AreEqual(1,f.Form.Designer.Controls.Count);
            Assert.AreEqual("Frame1",f.Form.Designer.Controls.Single().Name);
        }
        private sealed class HideControlsProvider : TypeDescriptionProvider
        {
            private readonly TypeDescriptionProvider parent; private int remaining=2;
            public HideControlsProvider(object owner) { parent=TypeDescriptor.GetProvider(owner); }
            public override ICustomTypeDescriptor GetTypeDescriptor(Type type,object instance) { return new HideControlsDescriptor(parent.GetTypeDescriptor(type,instance),()=>remaining-- > 0); }
        }
        private sealed class HideControlsDescriptor : CustomTypeDescriptor
        {
            private readonly Func<bool> visible; public HideControlsDescriptor(ICustomTypeDescriptor parent,Func<bool> visible):base(parent) { this.visible=visible; }
            public override PropertyDescriptorCollection GetProperties() { if(visible()) return base.GetProperties(); return new PropertyDescriptorCollection(base.GetProperties().Cast<PropertyDescriptor>().Where(p=>p.Name!="Controls").ToArray()); }
        }
        private static void ChangeGeometry(VbeFormsPartialTests.FakeControl c, string field, double value)
        {
            typeof(VbeFormsPartialTests.FakeControl).GetProperty(field).SetValue(c,value);
        }
        private static readonly object[][] InvalidBoxes = {
            new object[]{"Left",double.NaN},new object[]{"Top",double.PositiveInfinity},
            new object[]{"Width",double.NaN},new object[]{"Height",double.NegativeInfinity},
            new object[]{"Left",-1d},new object[]{"Top",-1d},new object[]{"Width",0d},new object[]{"Height",0d},
            new object[]{"Left",32768d},new object[]{"Top",32768d},new object[]{"Width",32768d},new object[]{"Height",32768d}
        };
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        private static void FramePreflightFailures(string route)
        {
            var f=FrameWithChildren(route);
            foreach(var path in new[]{null,"Controls/Frame1/Controls/Title"}) {
                var r=f.Request();r.ControlPath=path;
                Assert.ThrowsException<ArgumentException>(()=>CopyFrame(f,r,route));
            }
            foreach(var box in InvalidBoxes) {
                var invalid=FrameWithChildren(route);ChangeGeometry(invalid.Frame,(string)box[0],(double)box[1]);
                var error=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(invalid,invalid.Request(),route));
                StringAssert.Contains(error.Message,"geometry");
                Assert.AreEqual(1,invalid.Form.Designer.Controls.Count);
            }
            foreach(var box in InvalidBoxes) {
                var invalid=FrameWithChildren(route);ChangeGeometry(invalid.Frame.Controls.Item("Title"),(string)box[0],(double)box[1]);
                var error=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(invalid,invalid.Request(),route));
                StringAssert.Contains(error.Message,"geometry");
                Assert.AreEqual(1,invalid.Form.Designer.Controls.Count);
            }
            foreach(var size in new[]{double.NaN,0d,201d}) {
                var invalid=FrameWithChildren(route);invalid.Frame.Controls.Item("Title").Font.Size=size;
                var error=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(invalid,invalid.Request(),route));
                StringAssert.Contains(error.Message,"font");
                Assert.AreEqual(1,invalid.Form.Designer.Controls.Count);
            }
            var unnamed=FrameWithChildren(route);unnamed.Frame.Controls.Item("Title").Font.Name=" ";
            var fontError=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(unnamed,unnamed.Request(),route));
            StringAssert.Contains(fontError.Message,"font");
            var race=FrameWithChildren(route);var request=race.Request();int reads=0;
            race.Form.Designer.Controls.BeforeEnumeration=()=>{if(++reads==3) race.Frame.Caption="changed in preflight";};
            var changed=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(race,request,route));
            StringAssert.Contains(changed.Message,"hierarchy changed");
            Assert.AreEqual(1,race.Form.Designer.Controls.Count);
        }

        private static void FrameReadbackFailures(string route)
        {
            foreach(var field in new[]{"Caption","Children","Left","Top","Width","Height"}) {
                var f=FrameWithChildren(route);var r=f.Request();
                f.Form.Designer.Controls.ConfigureAdded=c=>{
                    if(field=="Caption") c.ReadOverrides["Caption"]="changed";
                    else if(field=="Children") c.Controls.CountOverride=99;
                    else c.ReadOverrides[field]=100d;
                };
                AssertRolledBack(f,r,route,route=="profiled"?"Copied Frame properties differ":"copied Frame did not retain");
            }
            foreach(var field in new[]{"Caption","BackColor","Font.Name","Font.Size","Font.Bold","Left","Top","Width","Height"}) {
                var f=FrameWithChildren(route);var r=f.Request();
                f.Form.Designer.Controls.ConfigureAdded=frame=>frame.Controls.ConfigureAdded=c=>{
                    if(c.Name.EndsWith("_Title",StringComparison.Ordinal)) {
                        if(field.StartsWith("Font.",StringComparison.Ordinal)) c.Font.ReadOverrides[field.Substring(5)]=field=="Font.Name"?(object)"Changed":field=="Font.Size"?(object)99d:true;
                        else c.ReadOverrides[field]=field=="Caption"?(object)"changed":field=="BackColor"?(object)123:100d;
                    }
                };
                AssertRolledBack(f,r,route,route=="labels"?"Copied Label differs":field=="Left"||field=="Top"||field=="Width"||field=="Height"?"geometry differs":"Copied Label differs");
            }
            var hidden=FrameWithChildren(route);var request=hidden.Request();hidden.Form.Designer.Controls.HideAdded=true;
            var hiddenError=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(hidden,request,route));
            StringAssert.Contains(hiddenError.Message,"rollback was incomplete");
            StringAssert.Contains(hiddenError.InnerException.Message,"not reflected");
            Assert.AreEqual(1,hidden.Form.Designer.Controls.Count);
            var extra=FrameWithChildren(route);request=extra.Request();
            extra.Form.Designer.Controls.ConfigureAdded=c=>c.Controls.AddExisting("Label","Unplanned");
            AssertRolledBack(extra,request,route,"not reflected");
            var missing=FrameWithChildren(route);request=missing.Request();
            missing.Form.Designer.Controls.ConfigureAdded=frame=>frame.Controls.ConfigureAdded=c=>c.Name="Unexpected"+c.Name;
            var absent=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(missing,request,route));
            StringAssert.Contains(absent.Message,"rollback");
            Assert.IsNotNull(absent.InnerException);
            StringAssert.Contains(absent.InnerException.Message,route=="profiled"?"Copied child missing":"absent");
            Assert.AreEqual(1,missing.Form.Designer.Controls.Count);
        }

        private static void FrameNativeAndRollbackFailures(string route)
        {
            var f=FrameWithChildren(route);var r=f.Request();f.Form.Designer.Controls.FailAdd=true;
            AssertRolledBack(f,r,route,"Native Add failed");
            f=FrameWithChildren(route);r=f.Request();f.Form.Designer.Controls.FailNextChildCaption=true;
            AssertRolledBack(f,r,route,"Caption setter failed");
            f=FrameWithChildren(route);r=f.Request();
            f.Form.Designer.Controls.ConfigureAdded=c=>{ c.Controls.FailNextCaption=true;c.Controls.FailNextRemove=true; };
            var childError=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(f,r,route));
            StringAssert.Contains(childError.Message,"rollback was incomplete");
            Assert.AreEqual(1,f.Form.Designer.Controls.Count);
            f=FrameWithChildren(route);r=f.Request();f.Form.Designer.Controls.FailNextCaption=true;f.Form.Designer.Controls.FailNextRemove=true;
            var rootError=Assert.ThrowsException<InvalidOperationException>(()=>CopyFrame(f,r,route));
            StringAssert.Contains(rootError.Message,"rollback was incomplete");
            Assert.AreEqual(2,f.Form.Designer.Controls.Count);
        }
    }
}
