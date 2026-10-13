using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie la frontière SIP VBA et les résultats de politique sans substituer la confiance d'un certificat.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeSignatureVerifierTests
    {
        /// <summary>Current Office uses its dedicated VBA subject for every previously supported file format.</summary>
        [DataTestMethod]
        [DataRow(".xlam"), DataRow(".xlsb"), DataRow(".xlsm"), DataRow(".xltm"), DataRow(".potm"), DataRow(".ppam"), DataRow(".ppsm"), DataRow(".pptm")]
        [DataRow(".vsdm"), DataRow(".vssm"), DataRow(".vstm"), DataRow(".docm"), DataRow(".dotm")]
        [DataRow(".xla"), DataRow(".xls"), DataRow(".xlt"), DataRow(".pot"), DataRow(".ppa"), DataRow(".pps"), DataRow(".ppt")]
        [DataRow(".mpp"), DataRow(".mpt"), DataRow(".pub"), DataRow(".vdw"), DataRow(".vsd"), DataRow(".vss"), DataRow(".vst"), DataRow(".doc"), DataRow(".dot"), DataRow(".wiz")]
        public void CurrentOfficeVbaSipKeepsExactFilePolicyAndProviderLifetime(string extension)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                string path = Path.Combine(fixture.Root, "macro" + extension); File.WriteAllText(path, "owned macro bytes");
                int verifies = 0, closes = 0;
                var verifier = new VbeSignatureVerifier { Provider = subject => throw new InvalidOperationException("An admitted current VBA provider must not probe a legacy subject."), ModernProvider = fixture.Provider };
                verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    Assert.AreEqual(new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE"), policy);
                    Assert.AreEqual(new IntPtr(-1), owner); Assert.AreEqual(2u, data.Ui); Assert.AreEqual(0x1080u, data.Flags);
                    var file = Marshal.PtrToStructure<VbeSignatureVerifier.TrustFile>(data.File);
                    Assert.AreEqual(path, file.Path); Assert.AreNotEqual(IntPtr.Zero, file.Handle);
                    Assert.AreEqual(OfficeVbaSignatureProvider.Subject, Marshal.PtrToStructure<Guid>(file.Subject));
                    Assert.AreEqual(0, fixture.FreeCalls, "Provider must outlive both verification and WinTrust state closure.");
                    if (data.Action == 2) { closes++; return 0; }
                    verifies++; return unchecked((int)0x800B0100);
                };
                dynamic result = verifier.Verify(path);
                Assert.IsTrue((bool)result.Available); Assert.AreEqual("NoSignature", (string)result.Status);
                Assert.AreEqual(OfficeVbaSignatureProvider.Subject.ToString("B"), (string)result.SipSubject);
                Assert.AreEqual(fixture.ModulePath, (string)result.Provider); Assert.IsFalse((bool)result.Trusted); Assert.IsNull((bool?)result.SignatureValid);
                Assert.AreEqual(1, verifies); Assert.AreEqual(1, closes); Assert.AreEqual(1, fixture.FreeCalls);
                File.AppendAllText(path, "file closed");
            }
        }

        /// <summary>DLL publisher metadata is read only after an actual successful Windows trust result.</summary>
        [DataTestMethod]
        [DataRow(0u, "Microsoft Corporation", true), DataRow(0u, "Other publisher", false), DataRow(0u, null, false)]
        [DataRow(5u, "Microsoft Corporation", false), DataRow(0x800B0109u, "Microsoft Corporation", false), DataRow(0x80092013u, "Microsoft Corporation", false)]
        public void MicrosoftModuleAdmissionRequiresPolicySuccessBeforePublisherIdentity(uint status, string signer, bool accepted)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                int verifies = 0, closes = 0, signerReads = 0;
                VbeSignatureVerifier.VerifyCall native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    Assert.AreEqual(new IntPtr(-1), owner); Assert.AreEqual(2u, data.Ui); Assert.AreEqual(0x1080u, data.Flags);
                    var file = Marshal.PtrToStructure<VbeSignatureVerifier.TrustFile>(data.File);
                    Assert.AreEqual(fixture.ModulePath, file.Path); Assert.AreEqual(IntPtr.Zero, file.Subject, "The DLL is checked as Authenticode, not as a macro.");
                    Assert.ThrowsException<IOException>(() => File.AppendAllText(fixture.ModulePath, "cannot replace a provider during validation"));
                    if (data.Action == 2) { closes++; return 0; }
                    verifies++; return unchecked((int)status);
                };
                Assert.AreEqual(accepted, VbeSignatureVerifier.VerifyProviderModule(fixture.ModulePath, native, path => { signerReads++; return signer; }));
                Assert.AreEqual(1, verifies); Assert.AreEqual(1, closes); Assert.AreEqual(status == 0 ? 1 : 0, signerReads);
                File.AppendAllText(fixture.ModulePath, "module check closed");
            }
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3)]
        public void ModulePolicyExceptionsCloseStateAndDoNotPromoteCertificatePresence(int scenario)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                int closes = 0, signerReads = 0;
                VbeSignatureVerifier.VerifyCall native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    if (data.Action == 2) { closes++; if (scenario == 1) throw new InvalidOperationException("closure failed"); return 0; }
                    if (scenario == 0) throw new InvalidOperationException("verification failed"); return 0;
                };
                Func<string, string> signer = path => { signerReads++; if (scenario == 2) throw new System.Security.Cryptography.CryptographicException("invalid certificate"); throw new InvalidOperationException("signer read failed"); };
                if (scenario == 2) Assert.IsFalse(VbeSignatureVerifier.VerifyProviderModule(fixture.ModulePath, native, signer));
                else Assert.ThrowsException<InvalidOperationException>(() => VbeSignatureVerifier.VerifyProviderModule(fixture.ModulePath, native, signer));
                Assert.AreEqual(1, closes); Assert.AreEqual(scenario < 2 ? 0 : 1, signerReads);
                File.AppendAllText(fixture.ModulePath, "no handle retained");
            }
        }

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void CurrentVbaProviderIsReleasedAfterVerificationOrStateClosureFailure(bool closeFailure)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                string path = Path.Combine(fixture.Root, "macro.xlsm"); File.WriteAllText(path, "owned macro bytes");
                int closes = 0;
                var verifier = new VbeSignatureVerifier { Provider = subject => null, ModernProvider = fixture.Provider };
                verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    if (data.Action == 2) { closes++; if (closeFailure) throw new InvalidOperationException("closure failed"); return 0; }
                    if (!closeFailure) throw new InvalidOperationException("verification failed"); return 0;
                };
                Assert.ThrowsException<InvalidOperationException>(() => verifier.Verify(path));
                Assert.AreEqual(1, closes); Assert.AreEqual(1, fixture.FreeCalls);
                File.AppendAllText(path, "macro file unlocked");
            }
        }

        /// <summary>An actual current-provider policy failure is terminal, never a legacy downgrade.</summary>
        [DataTestMethod]
        [DataRow(0u), DataRow(0x800B0100u), DataRow(0x80096010u), DataRow(0x80096004u), DataRow(0x800B0109u)]
        [DataRow(0x800B0101u), DataRow(0x800B010Cu), DataRow(0x80092013u), DataRow(0x800B010Eu), DataRow(5u)]
        public void CurrentProviderPolicyResultNeverFallsBackToAnOlderSip(uint code)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                string path = Path.Combine(fixture.Root, "macro.xlsm"); File.WriteAllText(path, "owned macro bytes");
                int verifies = 0, closes = 0;
                var verifier = new VbeSignatureVerifier
                {
                    ModernProvider = fixture.Provider,
                    Provider = subject => throw new InvalidOperationException("Policy failure cannot downgrade the provider.")
                };
                verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    Assert.AreEqual(OfficeVbaSignatureProvider.Subject, Marshal.PtrToStructure<Guid>(Marshal.PtrToStructure<VbeSignatureVerifier.TrustFile>(data.File).Subject));
                    if (data.Action == 2) { closes++; return 0; }
                    verifies++; return unchecked((int)code);
                };
                dynamic result = verifier.Verify(path);
                Assert.AreEqual(code == 0, (bool)result.Trusted);
                Assert.AreEqual(code == 0 ? true : code == 0x80096010u || code == 0x80096004u ? (bool?)false : null, (bool?)result.SignatureValid);
                Assert.AreEqual("0x" + code.ToString("X8"), (string)result.NativeStatus);
                Assert.AreEqual(1, verifies); Assert.AreEqual(1, closes); Assert.AreEqual(1, fixture.FreeCalls);
            }
        }

        /// <summary>A missing or rejected modern provider preserves the exact older per-format contract.</summary>
        [DataTestMethod]
        [DataRow(".xlsm", false), DataRow(".xls", false), DataRow(".xlsm", true), DataRow(".xls", true)]
        public void UnavailableCurrentProviderRetainsTheLegacyFormatBoundary(string extension, bool untrustedModule)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                if (untrustedModule) fixture.TrustResult = false;
                else fixture.Registration.Clear();
                string path = Path.Combine(fixture.Root, "macro" + extension); File.WriteAllText(path, "owned macro bytes");
                string provider = Path.Combine(fixture.Root, extension == ".xls" ? "msosip.dll" : "msosipx.dll"); File.WriteAllText(provider, "simulated legacy boundary");
                Guid expected = extension == ".xls" ? VbeSignatureVerifier.LegacySubject : VbeSignatureVerifier.XmlSubject;
                int legacyReads = 0, verifies = 0, closes = 0;
                var verifier = new VbeSignatureVerifier { ModernProvider = fixture.Provider, Provider = subject => { legacyReads++; Assert.AreEqual(expected, subject); return provider; } };
                verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                {
                    Assert.AreEqual(expected, Marshal.PtrToStructure<Guid>(Marshal.PtrToStructure<VbeSignatureVerifier.TrustFile>(data.File).Subject));
                    if (data.Action == 2) { closes++; return 0; }
                    verifies++; return unchecked((int)0x800B0100);
                };
                dynamic result = verifier.Verify(path);
                Assert.AreEqual("NoSignature", (string)result.Status); Assert.AreEqual(provider, (string)result.Provider);
                Assert.AreEqual(1, legacyReads); Assert.AreEqual(1, verifies); Assert.AreEqual(1, closes);
                Assert.AreEqual(0, fixture.LoadCalls); Assert.AreEqual(0, fixture.FreeCalls);
            }
        }

        /// <summary>A successful policy result cannot hide an uncertain provider cleanup.</summary>
        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void CurrentProviderReleaseFailureCannotReturnATrustedResult(bool throwing)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.FreeResult = false;
                if (throwing) fixture.FreeFailure = new InvalidOperationException("uncertain release");
                string path = Path.Combine(fixture.Root, "macro.xlsm"); File.WriteAllText(path, "owned macro bytes");
                int closes = 0;
                var verifier = new VbeSignatureVerifier { ModernProvider = fixture.Provider, Provider = subject => throw new InvalidOperationException("No downgrade.") };
                verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) => { if (data.Action == 2) closes++; return 0; };
                Assert.ThrowsException<InvalidOperationException>(() => verifier.Verify(path));
                Assert.AreEqual(1, closes); Assert.AreEqual(1, fixture.FreeCalls);
                File.AppendAllText(path, "macro file lock released"); File.AppendAllText(fixture.ModulePath, "module file lock released");
            }
        }

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
                    var verifier = new VbeSignatureVerifier { Provider = subject => { selected = subject; return null; }, ModernProvider = null };
                    dynamic result = verifier.Verify(path);
                    Assert.IsFalse((bool)result.Available); Assert.IsNull((bool?)result.Trusted); Assert.IsNull((bool?)result.SignatureValid);
                    Assert.AreEqual(64, ((string)result.FileSha256).Length);
                    Assert.AreEqual(extension.EndsWith("m") || new[] { ".xlsb", ".vsdm", ".vssm", ".vstm" }.Contains(extension) ? VbeSignatureVerifier.XmlSubject : VbeSignatureVerifier.LegacySubject, selected);
                }
                string macro = Path.Combine(folder, "test.xlsm");
                foreach (string provider in new[] { " ", "msosipx.dll", Path.Combine(folder, "wrong.dll"), Path.Combine(folder, "msosipx.dll") })
                {
                    if (provider.EndsWith("wrong.dll")) File.WriteAllText(provider, "not a SIP");
                    var verifier = new VbeSignatureVerifier { Provider = subject => provider, ModernProvider = null };
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
                    var verifier = new VbeSignatureVerifier { Provider = subject => provider, ModernProvider = null };
                    verifier.Native = (IntPtr owner, ref Guid policy, ref VbeSignatureVerifier.TrustData data) =>
                    {
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
