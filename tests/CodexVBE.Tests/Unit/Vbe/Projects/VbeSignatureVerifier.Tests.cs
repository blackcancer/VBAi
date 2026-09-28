using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie la frontière SIP VBA et les résultats de politique sans substituer la confiance d'un certificat.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeSignatureVerifierTests
    {
        /// <summary>Refuse formats, chemins et tailles hors contrat.</summary>
        [TestMethod]
        public void SignatureFileGuardsRejectUnsupportedPathsAndEmptyOrExcessiveFiles()
        {
            var verifier = new VbeSignatureVerifier();
            foreach (string path in new[] { null, " ", "relative.xlsm" })
                Assert.ThrowsException<ArgumentException>(() => verifier.Verify(path));
            Assert.ThrowsException<NotSupportedException>(() => verifier.Verify(Path.Combine(Path.GetTempPath(), "test.swp")));
            Assert.ThrowsException<FileNotFoundException>(() => verifier.Verify(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsm")));
            string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsm");
            try
            {
                using (var stream = File.Create(file)) { }
                Assert.ThrowsException<InvalidOperationException>(() => verifier.Verify(file));
                using (var stream = File.OpenWrite(file)) stream.SetLength(512L * 1024 * 1024 + 1);
                Assert.ThrowsException<InvalidOperationException>(() => verifier.Verify(file));
            }
            finally { File.Delete(file); }
        }

        /// <summary>Exerce chaque extension documentée et toutes les absences de SIP sans appeler WinTrust.</summary>
        [TestMethod]
        public void SupportedFormatsSelectOnlyTheirExactMicrosoftVbaSip()
        {
            string folder = Path.Combine(Path.GetTempPath(), "VBAi-sip-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            try
            {
                foreach (string extension in new[] { ".xlam", ".xlsb", ".xlsm", ".xltm", ".potm", ".ppam", ".ppsm", ".pptm", ".vsdm", ".vssm", ".vstm", ".docm", ".dotm",
                    ".xla", ".xls", ".xlt", ".pot", ".ppa", ".pps", ".ppt", ".mpp", ".mpt", ".pub", ".vdw", ".vsd", ".vss", ".vst", ".doc", ".dot", ".wiz" })
                {
                    string path = Path.Combine(folder, "test" + extension); File.WriteAllText(path, "document");
                    Guid selected = Guid.Empty;
                    var verifier = new VbeSignatureVerifier { Provider = subject => { selected = subject; return null; } };
                    dynamic result = verifier.Verify(path);
                    Assert.IsFalse((bool)result.Available); Assert.IsNull((bool?)result.Trusted); Assert.IsNull((bool?)result.SignatureValid);
                    Assert.AreEqual(64, ((string)result.FileSha256).Length);
                    Assert.AreEqual(extension.EndsWith("m") || new[] { ".xlsb", ".vsdm", ".vssm", ".vstm" }.Contains(extension) ? VbeSignatureVerifier.XmlSubject : VbeSignatureVerifier.LegacySubject, selected);
                }
                string macro = Path.Combine(folder, "test.xlsm");
                foreach (string provider in new[] { " ", "msosipx.dll", Path.Combine(folder, "wrong.dll"), Path.Combine(folder, "msosipx.dll") })
                {
                    if (provider.EndsWith("wrong.dll")) File.WriteAllText(provider, "not a SIP");
                    var verifier = new VbeSignatureVerifier { Provider = subject => provider };
                    Assert.IsFalse((bool)((dynamic)verifier.Verify(macro)).Available);
                }
                // Read the actual registered-provider boundary as well; absence is a legitimate native result.
                var actual = new VbeSignatureVerifier(); actual.Provider(VbeSignatureVerifier.XmlSubject);
                actual.Provider(Guid.NewGuid());
                Assert.IsNotNull(actual.Provider(new Guid("9BA61D3F-E73A-11D0-8CD2-00C04FC295EE")), "The standard Windows SIP registry must remain readable.");
            }
            finally { Directory.Delete(folder, true); }
        }

        /// <summary>Exerce les HRESULT, structures et libérations même quand la frontière native lève une exception.</summary>
        [TestMethod]
        public void WindowsPolicyResultsAndNativeLifetimeRemainDistinctFromCertificatePresence()
        {
            string folder = Path.Combine(Path.GetTempPath(), "VBAi-sip-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "macro.xlsm"), provider = Path.Combine(folder, "msosipx.dll");
            File.WriteAllText(path, "macro bytes"); File.WriteAllText(provider, "fixture native boundary");
            try
            {
                uint[] codes = { 0, 0x800B0100, 0x80096010, 0x80096004, 0x800B0109, 0x800B0101, 0x800B010C, 0x80092013, 0x800B010E, 5 };
                string[] statuses = { "Trusted", "NoSignature", "BadDigest", "InvalidSignature", "UntrustedRoot", "ExpiredCertificate", "RevokedCertificate", "RevocationIndeterminate", "RevocationIndeterminate", "VerificationFailed" };
                for (int scenario = 0; scenario < codes.Length + 2; scenario++)
                {
                    int selected = scenario, verifies = 0, closes = 0;
                    var verifier = new VbeSignatureVerifier { Provider = subject => provider };
                    verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) => {
                        Assert.AreEqual(new IntPtr(-1), owner); Assert.AreEqual(2u, data.Ui); Assert.AreEqual(1u, data.Choice);
                        Assert.AreEqual(0x1080u, data.Flags); Assert.AreEqual((uint)Marshal.SizeOf(typeof(VbeSignatureVerifier.TrustData)), data.Size);
                        var file = Marshal.PtrToStructure<VbeSignatureVerifier.TrustFile>(data.File);
                        Assert.AreEqual(path, file.Path); Assert.AreNotEqual(IntPtr.Zero, file.Handle);
                        Assert.AreEqual(VbeSignatureVerifier.XmlSubject, Marshal.PtrToStructure<Guid>(file.Subject));
                        if (data.Action == 2) { closes++; if (selected == codes.Length + 1) throw new InvalidOperationException("close failed"); return 0; }
                        verifies++; data.State = new IntPtr(7);
                        if (selected >= codes.Length) throw new InvalidOperationException("verify failed");
                        return unchecked((int)codes[selected]);
                    };
                    if (scenario >= codes.Length) Assert.ThrowsException<InvalidOperationException>(() => verifier.Verify(path));
                    else
                    {
                        dynamic result = verifier.Verify(path);
                        Assert.AreEqual(statuses[scenario], (string)result.Status);
                        Assert.AreEqual(scenario == 0, (bool)result.Trusted);
                        bool? expected = scenario == 0 ? true : scenario == 2 || scenario == 3 ? (bool?)false : null;
                        Assert.AreEqual(expected, (bool?)result.SignatureValid);
                        Assert.AreEqual("0x" + codes[scenario].ToString("X8"), (string)result.NativeStatus);
                        Assert.IsFalse((bool)result.NetworkRetrieval);
                    }
                    Assert.AreEqual(1, verifies); Assert.AreEqual(1, closes);
                }
                File.AppendAllText(path, "changed after verification"); // Every retained file handle and native buffer was released.
            }
            finally { Directory.Delete(folder, true); }
        }
    }
}
