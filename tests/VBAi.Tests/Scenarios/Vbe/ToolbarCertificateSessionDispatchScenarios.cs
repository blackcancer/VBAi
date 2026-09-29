namespace VBAi.Tests.Unit
{
    using System;
    using System.Security.Cryptography;
    using System.Security.Cryptography.X509Certificates;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class ToolbarCustomizationTests
    {
        [TestMethod]
        public void SessionRoutesEveryToolbarCustomizationCommandThroughItsOwnNativeContracts()
        {
            var host=new Host();var bar=host.CommandBars.Add("Standard",1,false,true);bar.Controls.Add(1,42,Type.Missing,1,false);var session=new VbeSession(host);
            Func<Request,string,object> run=(request,command)=>{request.Command=command;var result=session.Execute(request);Assert.IsTrue(result.Ok,result.Error);return result.Data;};
            var r=new Request{ObjectName="Standard"};r.ExpectedToolbarControlsVersion=(string)Data(run(r,"toolbar_controls"))["ToolbarControlsVersion"];r.ControlId=42;r.ControlCaption="Native command";
            Assert.IsTrue((bool)Data(run(r,"add_toolbar_command"))["Verified"]);Assert.AreEqual(2,bar.Controls.Count);
            r.ExpectedToolbarControlsVersion=(string)Data(run(r,"toolbar_controls"))["ToolbarControlsVersion"];r.InsertIndex=2;
            Assert.IsTrue((bool)Data(run(r,"remove_toolbar_command"))["Verified"]);Assert.AreEqual(1,bar.Controls.Count);
            r=new Request{ObjectName="Dispatch",ExpectedToolbarCollectionVersion=(string)Data(run(new Request(),"list_toolbars"))["ToolbarCollectionVersion"]};
            var created=Data(run(r,"create_toolbar"));Assert.IsTrue((bool)created["Verified"]);
            r.ObjectName=(string)created["ObjectName"];r.ExpectedToolbarCollectionVersion=(string)created["ToolbarCollectionVersion"];r.ExpectedToolbarControlsVersion=(string)Data(run(r,"toolbar_controls"))["ToolbarControlsVersion"];
            Assert.IsTrue((bool)Data(run(r,"remove_toolbar"))["Verified"]);Assert.AreEqual(1,host.CommandBars.Count);
        }
    }
    [TestClass,TestCategory("Unit")]
    public sealed class CertificateSessionDispatchTests
    {
        [TestMethod]
        public void SessionCertificateTrustRequiresExactCurrentUserMatchAndDisposesReadOnlyStore()
        {
            using(var rsa=RSA.Create(2048))
            using(var certificate=new CertificateRequest("CN=isolated dispatch fixture",rsa,HashAlgorithmName.SHA256,RSASignaturePadding.Pkcs1).CreateSelfSigned(DateTimeOffset.UtcNow.AddHours(-1),DateTimeOffset.UtcNow.AddHours(1)))
            {
                var session=new VbeSession(new object());var store=new CertificateDispatchStore();
                session.SigningStore=location=>{Assert.AreEqual(StoreLocation.CurrentUser,location);return store;};
                foreach(string invalid in new[]{null,"bad",new string('z',40)})Assert.ThrowsException<ArgumentException>(()=>session.Execute(new Request{Command="certificate_trust",CertificateThumbprint=invalid}));
                Assert.IsNull(store.Opened);
                foreach(int count in new[]{0,2,1})
                {
                    store=new CertificateDispatchStore();for(int i=0;i<count;i++)store.Certificates.Add(certificate);
                    var request=new Request{Command="certificate_trust",CertificateThumbprint=certificate.Thumbprint.ToLowerInvariant()};
                    if(count!=1)Assert.ThrowsException<InvalidOperationException>(()=>session.Execute(request));
                    else {var response=session.Execute(request);Assert.IsTrue(response.Ok);dynamic result=response.Data;Assert.IsTrue((bool)result.OfflineOnly);Assert.IsFalse((bool)result.MacroSignatureVerified);}
                    Assert.IsTrue(store.Disposed);Assert.AreEqual(OpenFlags.ReadOnly,store.Opened.Value);
                }
            }
        }
    }
}
