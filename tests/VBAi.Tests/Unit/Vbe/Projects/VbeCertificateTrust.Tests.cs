namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography;
    using System.Security.Cryptography.X509Certificates;
    using System.Web.Script.Serialization;
    using System.Runtime.InteropServices;
    using System.ComponentModel;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class CertificateTrustTests
    {
        [TestMethod]
        public void NativeOfflinePolicyMatrixPreservesFlagsAndDistinguishesTrustFromMissingEvidence()
        {
            var chain=VbeCertificateTrust.ReadChain;var policy=VbeCertificateTrust.ReadPolicy;var release=VbeCertificateTrust.ReleaseChain;var allocate=VbeCertificateTrust.AllocatePointers;
            using(var rsa=RSA.Create(2048))
            using(var certificate=new CertificateRequest("CN=isolated native policy",rsa,HashAlgorithmName.SHA256,RSASignaturePadding.Pkcs1).CreateSelfSigned(DateTimeOffset.UtcNow.AddHours(-1),DateTimeOffset.UtcNow.AddHours(1)))
            try
            {
                uint errors=0,policyError=0;bool chainOk=true,policyOk=true;int freed=0;
                VbeCertificateTrust.ReadChain=(IntPtr engine,IntPtr cert,IntPtr time,IntPtr store,ref VbeCertificateTrust.ChainParameters parameters,uint flags,IntPtr reserved,out IntPtr pointer)=>{
                    Assert.AreEqual(certificate.Handle,cert);Assert.AreEqual(0xC0000104u,flags);Assert.AreEqual(1u,parameters.Usage.Usage.Count);
                    Assert.AreEqual("1.3.6.1.5.5.7.3.3",Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(parameters.Usage.Usage.Identifiers)));
                    pointer=IntPtr.Zero;if(!chainOk)return false;
                    pointer=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(VbeCertificateTrust.ChainHeader)));
                    Marshal.StructureToPtr(new VbeCertificateTrust.ChainHeader{Size=Marshal.SizeOf(typeof(VbeCertificateTrust.ChainHeader)),ErrorStatus=errors},pointer,false);return true;
                };
                VbeCertificateTrust.ReadPolicy=(IntPtr id,IntPtr pointer,ref VbeCertificateTrust.PolicyParameters parameters,ref VbeCertificateTrust.PolicyStatus status)=>{Assert.AreEqual(new IntPtr(2),id);status.Error=policyError;status.ChainIndex=2;status.ElementIndex=3;return policyOk;};
                VbeCertificateTrust.ReleaseChain=pointer=>{freed++;Marshal.FreeHGlobal(pointer);};
                foreach(var pair in new[]{Tuple.Create(0u,0u,"Trusted"),Tuple.Create(0u,1u,"NotTrusted"),Tuple.Create(0x40u,0u,"IndeterminateOffline"),Tuple.Create(0x1000000u,0x80092013u,"IndeterminateOffline"),Tuple.Create(0x40u,0x800B010Eu,"IndeterminateOffline"),Tuple.Create(0x40u,1u,"NotTrusted"),Tuple.Create(0x20u,0u,"NotTrusted")})
                {errors=pair.Item1;policyError=pair.Item2;dynamic result=VbeCertificateTrust.Evaluate(certificate);Assert.AreEqual(pair.Item3,(string)result.State);Assert.AreEqual(pair.Item3=="Trusted",(bool)result.Trusted);Assert.IsTrue((bool)result.OfflineOnly);Assert.IsFalse((bool)result.PrivateKeyAccessed);Assert.AreEqual(2,(int)result.ChainIndex);}
                Assert.AreEqual(7,freed);chainOk=false;Assert.ThrowsException<Win32Exception>(()=>VbeCertificateTrust.Evaluate(certificate));Assert.AreEqual(7,freed);
                chainOk=true;policyOk=false;Assert.ThrowsException<Win32Exception>(()=>VbeCertificateTrust.Evaluate(certificate));Assert.AreEqual(8,freed);
                VbeCertificateTrust.AllocatePointers=size=>{throw new OutOfMemoryException("native allocation rejected");};
                Assert.ThrowsException<OutOfMemoryException>(()=>VbeCertificateTrust.Evaluate(certificate));Assert.AreEqual(8,freed);
            }
            finally{VbeCertificateTrust.ReadChain=chain;VbeCertificateTrust.ReadPolicy=policy;VbeCertificateTrust.ReleaseChain=release;VbeCertificateTrust.AllocatePointers=allocate;}
        }
        [TestMethod]
        public void UninstalledSelfSignedCertificateCannotBecomeTrustedOffline()
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest("CN=VBAi offline trust fixture", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                var usages = new OidCollection { new Oid("1.3.6.1.5.5.7.3.3") };
                request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(usages, true));
                using (var certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddHours(-1), DateTimeOffset.UtcNow.AddDays(1)))
                {
                    var json = new JavaScriptSerializer();
                    var result = json.Deserialize<Dictionary<string, object>>(json.Serialize(VbeCertificateTrust.Evaluate(certificate)));
                    Assert.AreEqual("NotTrusted", result["State"]);
                    Assert.AreEqual(false, result["Trusted"]); Assert.AreEqual(true, result["OfflineOnly"]);
                    Assert.AreEqual(false, result["MacroSignatureVerified"]); Assert.AreEqual(false, result["PrivateKeyAccessed"]);
                }
            }
            Assert.ThrowsException<ArgumentNullException>(() => VbeCertificateTrust.Evaluate(null));
            var session = new VbeSession(new object());
            Assert.ThrowsException<ArgumentException>(() => session.Execute(new Request { Command = "certificate_trust", CertificateThumbprint = "invalid" }));
        }
    }
}
