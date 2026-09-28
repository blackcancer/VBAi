namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsInitializerTests
    {
        private static Request Binding(Fixture f)
        {
            var request=RequestFor(f);request.ExpectedHostPath=@"C:\Tests\Macro.xlsm";
            request.SheetName="Sheet1";request.RangeAddress="A1";return request;
        }
        [TestMethod]
        public void BindingRequiresSavedHostAndEveryRevisionBeforeWriting()
        {
            foreach(int scenario in new[]{0,1,2,3})
            {
                var f=Create();f.Project.FileName=@"C:\Tests\Macro.xlsm";var request=Binding(f);
                if(scenario==0)request.ExpectedHostPath=null;if(scenario==1)request.ExpectedHostPath="relative.xlsm";
                if(scenario==2)request.ExpectedTreeVersion=null;if(scenario==3)request.ExpectedSha256=null;
                Assert.ThrowsException<ArgumentException>(()=>f.Service.SetListBinding(request));Assert.AreEqual(0,f.Form.CodeModule.InsertCount);
            }
            foreach(string path in new[]{" ","relative.xlsm",@"C:\Tests\Other.xlsm",@"C:\Tests\Macro.txt"})
            {
                var f=Create();f.Project.FileName=path;var request=Binding(f);
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.SetListBinding(request));Assert.AreEqual(0,f.Form.CodeModule.InsertCount);
            }
        }
        [TestMethod]
        public void BindingRefusesStaleTreeMissingControlWrongTypeAndChangedCode()
        {
            foreach(int scenario in new[]{0,1,2,3})
            {
                var f=Create();f.Project.FileName=@"C:\Tests\Macro.xlsm";
                if(scenario==2)f.Form.Designer.Controls.Add("Label","Label1");
                var request=Binding(f);
                if(scenario==0)request.ExpectedTreeVersion="stale";
                if(scenario==1)request.ControlPath="Controls/Missing";
                if(scenario==2)request.ControlPath="Controls/Label1";
                if(scenario==3)request.ExpectedSha256="stale";
                if(scenario==1)Assert.ThrowsException<ArgumentException>(()=>f.Service.SetListBinding(request));
                else Assert.ThrowsException<InvalidOperationException>(()=>f.Service.SetListBinding(request));
                Assert.AreEqual(0,f.Form.CodeModule.InsertCount);
            }
            var valid=Create();valid.Project.FileName=@"C:\Tests\Macro.xlsm";valid.Form.Designer.Controls.Add("ListBox","ListBox1");
            var r=Binding(valid);r.ControlPath="Controls/ListBox1";
            Assert.IsTrue((bool)((dynamic)valid.Service.SetListBinding(r)).Verified);
        }
        [TestMethod]
        public void ExcelBindingNamesAndRangesCoverLiteralAndIndependentExcelBounds()
        {
            foreach(string sheet in new[]{(string)null,new string('a',32),"'Sheet","Sheet'","Sheet\n","Sheet/one"})
                Assert.ThrowsException<ArgumentException>(()=>VbeForms.ExcelBindingColumns(sheet,"A1"));
            foreach(string address in new[]{(string)null,"A1:XFE1","A1:A1048577","A2:A1"})
                Assert.ThrowsException<ArgumentException>(()=>VbeForms.ExcelBindingColumns("Sheet",address));
            Assert.AreEqual(1,VbeForms.ExcelBindingColumns("Sheet","a1"));
        }
    }
}
