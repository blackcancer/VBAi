namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography;
    using System.Security.Cryptography.X509Certificates;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass, TestCategory("Unit")]
    public sealed class CertificateTrustTests
    {
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
