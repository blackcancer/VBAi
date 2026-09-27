namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Security.Cryptography;
    using System.Security.Cryptography.X509Certificates;
    using CodexVBE;

    public sealed partial class VbeSessionCoverageTests
    {
        public sealed class SessionHost
        {
            public string Version => "7.1";
            public List<VbeSessionTests.FakeProject> VBProjects { get; } = new List<VbeSessionTests.FakeProject>();
            public object ActiveVBProject { get; set; }
            public object ActiveWindow { get; set; }
            public object ActiveCodePane { get; set; }
            public List<object> Windows { get; } = new List<object>();
            public List<object> CodePanes { get; } = new List<object>();
            public List<object> AddIns { get; } = new List<object>();
            public List<object> CommandBars { get; } = new List<object>();
        }

        public sealed class SigningWorkbook
        {
            public string FullName { get; set; }
            public bool VBASigned { get; set; }
        }

        public sealed class SigningExcel
        {
            public int Hwnd => 100;
            public List<SigningWorkbook> Workbooks { get; } = new List<SigningWorkbook>();
        }

        private sealed class SigningHost : VbeProjectComponents.IExcelHostProbe
        {
            public bool IsExcel { get; set; }
            public int CurrentProcessId => 42;
            public SigningExcel Excel { get; } = new SigningExcel();
            public object ExcelApplication() { return Excel; }
            public uint WindowProcessId(IntPtr window) { return 42; }
        }

        private sealed class SigningStore : VbeSession.ISigningStore
        {
            public X509Certificate2Collection Certificates { get; } = new X509Certificate2Collection();
            public bool Disposed { get; private set; }
            public OpenFlags? Flags { get; private set; }
            public Exception OpenError { get; set; }
            public void Open(OpenFlags flags)
            {
                Flags = flags;
                if (OpenError != null) throw OpenError;
            }
            public void Dispose() { Disposed = true; }
        }

        private sealed class SigningFixture : IDisposable
        {
            public readonly DateTime Now = new DateTime(2030, 1, 10, 12, 0, 0, DateTimeKind.Local);
            public readonly SessionHost Vbe = new SessionHost();
            public readonly VbeSessionTests.FakeProject Project = new VbeSessionTests.FakeProject { Name = "SigningProject", Mode = 2, Saved = true };
            public readonly SigningHost Host = new SigningHost();
            public readonly SigningWorkbook Workbook = new SigningWorkbook();
            public readonly SigningStore Personal = new SigningStore();
            public readonly SigningStore Machine = new SigningStore();
            public readonly List<StoreLocation> Opened = new List<StoreLocation>();
            public readonly VbeSession Session;
            public readonly string FilePath;
            public int Scheduled;
            public Request ScheduledRequest;
            public readonly object NativeResult = new object();

            public SigningFixture(bool excel = false)
            {
                FilePath = Path.Combine(Path.GetTempPath(), "CodexVBE-Signing-" + Guid.NewGuid().ToString("N") + ".xlsm");
                File.WriteAllText(FilePath, "disposable signing path fixture");
                SetPath(FilePath);
                Host.IsExcel = excel;
                Host.Excel.Workbooks.Add(Workbook);
                Vbe.VBProjects.Add(Project);
                Vbe.ActiveVBProject = Project;
                Session = new VbeSession(Vbe, Host);
                Session.SigningProcessName = () => excel ? "EXCEL" : "SLDWORKS";
                Session.SigningClock = () => Now;
                Session.SigningStore = location => { Opened.Add(location); return location == StoreLocation.CurrentUser ? Personal : Machine; };
                Session.SignatureScheduler = request => { Scheduled++; ScheduledRequest = request; return NativeResult; };
            }

            public void SetPath(string path) { Project.FileName = path; Workbook.FullName = path; }
            public Request Request(string thumbprint)
            {
                dynamic state = Session.Execute(new Request { Command = "project_properties", Project = Project.Name }).Data;
                return new Request { Command = "sign_project", Project = Project.Name, ExpectedMode = 2,
                    ExpectedProjectVersion = state.Version, CertificateThumbprint = thumbprint };
            }
            public void Dispose() { File.Delete(FilePath); }
        }

        private sealed class CertificateFixture : IDisposable
        {
            private readonly RSA key = RSA.Create();
            public X509Certificate2 Certificate { get; }
            public CertificateFixture(string name = "CodexVBE Unit Signing", bool signing = true,
                DateTimeOffset? before = null, DateTimeOffset? after = null)
            {
                key.KeySize = 1024;
                var request = new CertificateRequest(new X500DistinguishedName(name.Length == 0 ? "" : "CN=" + name), key,
                    HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                if (signing) request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection {
                    new Oid("1.3.6.1.5.5.7.3.1"), new Oid("1.3.6.1.5.5.7.3.3") }, false));
                Certificate = request.CreateSelfSigned(before ?? new DateTimeOffset(2029, 1, 1, 0, 0, 0, TimeSpan.Zero),
                    after ?? new DateTimeOffset(2031, 1, 1, 0, 0, 0, TimeSpan.Zero));
            }
            public void Dispose() { Certificate.Dispose(); key.Dispose(); }
        }
    }
}
