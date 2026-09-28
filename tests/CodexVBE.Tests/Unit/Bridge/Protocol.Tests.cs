namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class ProtocolTests
    {
        [TestMethod]
        public void RequestRoundtripPreservesCommandScopePreconditionsSelectionAndTypedDesignerValues()
        {
            var original=new Request {Command="replace_lines",Project="P",Module="M",StartLine=2,StartColumn=3,EndLine=4,EndColumn=5,Count=6,
                ExpectedSha256="sha",Text="été",Items=new[] {"a","b"},Query="needle",Action="replace",ControlId=7,ControlCaption="control",WindowCaption="window",WindowType=3,
                ExpectedMode=2,Form="F",Control="C",ControlType="Label",Left=1.5,Top=2.5,Width=3.5,Height=4.5,Caption="title",ExpectedFormVersion="form-version",Property="Caption",Value=true,
                Path=@"C:\Temp\fixture",SourceEncoding="utf-8",Guid="guid",Major=1,Minor=2,Offset=3,Limit=4,RowIndex=5,ExpectedListVersion="list",TypeIndex=6,TypeIdentity="type",ExpectedReferencesVersion="references",
                ParentPath="parent",ExpectedTreeVersion="tree",ControlPath="control-path",ExpectedProjectVersion="project",ExpectedHostPath="host",CertificateThumbprint="thumbprint",ExpectedComponentVersion="component",
                NewName="new-name",Procedure="procedure",EventName="event",Expression="expression",NewExpression="new-expression",Context="context",WatchType="expression",Diagnostic="diagnostic",Button="button",Pane="locals",
                PathSegments=new[] {"one","two"},ObjectName="object",ProcKind=3,InsertIndex=8,ZPosition=1,WholeWord=true,MatchCase=true,PatternSearch=true,IncludeCallStack=true,FontName="Consolas",FontSize=11.5,FontBold=true};
            var json=new JavaScriptSerializer();var read=json.Deserialize<Request>(json.Serialize(original));
            Assert.AreEqual("replace_lines",read.Command);Assert.AreEqual("P",read.Project);Assert.AreEqual("M",read.Module);Assert.AreEqual("sha",read.ExpectedSha256);
            Assert.AreEqual("été",read.Text);CollectionAssert.AreEqual(original.Items,read.Items);CollectionAssert.AreEqual(original.PathSegments,read.PathSegments);
            Assert.AreEqual(2,read.StartLine);Assert.AreEqual(5,read.EndColumn);Assert.AreEqual(5,read.RowIndex);Assert.AreEqual(8,read.InsertIndex);
            Assert.AreEqual(true,read.Value);Assert.AreEqual(1.5,read.Left);Assert.AreEqual(11.5,read.FontSize);Assert.IsTrue(read.FontBold);
            Assert.AreEqual("tree",read.ExpectedTreeVersion);Assert.AreEqual("host",read.ExpectedHostPath);Assert.AreEqual("thumbprint",read.CertificateThumbprint);
            Assert.IsTrue(read.IncludeCallStack);Assert.IsTrue(read.WholeWord);Assert.IsTrue(read.MatchCase);Assert.IsTrue(read.PatternSearch);
            // JSON object member order is not part of the bridge protocol.
            // Coverage instrumentation can change reflection enumeration order.
            var before = (IDictionary<string, object>)json.DeserializeObject(json.Serialize(original));
            var after = (IDictionary<string, object>)json.DeserializeObject(json.Serialize(read));
            Assert.AreEqual(before.Count, after.Count);
            foreach (var member in before)
            {
                Assert.IsTrue(after.ContainsKey(member.Key), member.Key);
                Assert.AreEqual(json.Serialize(member.Value), json.Serialize(after[member.Key]), member.Key);
            }
        }

        [TestMethod]
        public void ResponseSuccessAndFailureKeepPayloadAndErrorStatesAcrossJson()
        {
            var json=new JavaScriptSerializer();var success=json.Deserialize<Response>(json.Serialize(Response.Success("payload")));
            Assert.IsTrue(success.Ok);Assert.AreEqual("payload",success.Data);Assert.IsNull(success.Error);
            var failure=json.Deserialize<Response>(json.Serialize(Response.Failure("declined")));
            Assert.IsFalse(failure.Ok);Assert.AreEqual("declined",failure.Error);Assert.IsNull(failure.Data);
        }
    }
}
