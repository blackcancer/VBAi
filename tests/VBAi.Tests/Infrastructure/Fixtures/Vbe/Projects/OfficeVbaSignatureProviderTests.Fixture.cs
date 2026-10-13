using System;
using System.Collections.Generic;
using System.IO;

namespace VBAi.Tests.Unit
{
    /// <summary>Owns a temporary module file and simulated native acquisitions; never loads a test DLL.</summary>
    internal sealed class OfficeVbaSignatureProviderFixture : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "VBAi-provider-" + Guid.NewGuid().ToString("N"));
        internal readonly Dictionary<string, string> Registration = new Dictionary<string, string>();
        internal readonly List<string> ExportNames = new List<string>();
        internal readonly OfficeVbaSignatureProvider Provider = new OfficeVbaSignatureProvider();
        internal string ModulePath, LoadedPath, InstallationRoot, MissingExport, LoadedArgument;
        internal uint LoadedFlags;
        internal int TrustCalls, LoadCalls, FreeCalls;
        internal bool TrustResult = true, FreeResult = true;
        internal IntPtr LoadResult = new IntPtr(123), FreedHandle;
        internal Exception TrustFailure, LoadFailure, ExportFailure, FreeFailure;

        internal OfficeVbaSignatureProviderFixture()
        {
            Directory.CreateDirectory(Root);
            ModulePath = Path.Combine(Root, "mso.dll");
            File.WriteAllText(ModulePath, "simulated native module boundary; not executable");
            LoadedPath = ModulePath;
            var sections = new[] { "CryptSIPDllVerifyIndirectData", "CryptSIPDllGetSignedDataMsg", "CryptSIPDllCreateIndirectData" };
            var functions = new[] { "MsoVBADigSigVerifyIndirectData", "MsoVBADigSigGetSignedDataMsg", "MsoVBADigSigCreateIndirectData" };
            for (int index = 0; index < sections.Length; index++)
            { Registration[sections[index] + "/Dll"] = "mso.dll"; Registration[sections[index] + "/FuncName"] = functions[index]; }
            Provider.ReadRegistration = (section, name) => Registration.TryGetValue(section + "/" + name, out string value) ? value : null;
            Provider.ReadLoadedPath = () => LoadedPath;
            Provider.ReadInstallationRoot = () => InstallationRoot;
            Provider.VerifyMicrosoftModule = path => { TrustCalls++; if (TrustFailure != null) throw TrustFailure; return TrustResult; };
            Provider.Load = (path, flags) => { LoadCalls++; LoadedArgument = path; LoadedFlags = flags; if (LoadFailure != null) throw LoadFailure; return LoadResult; };
            Provider.ReadAcquiredPath = handle => ModulePath;
            Provider.Export = (handle, name) => { ExportNames.Add(name); if (ExportFailure != null) throw ExportFailure; return name == MissingExport ? IntPtr.Zero : new IntPtr(124); };
            Provider.Free = handle => { FreeCalls++; FreedHandle = handle; if (FreeFailure != null) throw FreeFailure; return FreeResult; };
        }
        public void Dispose()
        {
            string prefix = Path.Combine(Path.GetTempPath(), "VBAi-provider-");
            if (!Path.GetFullPath(Root).StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Owned fixture directory changed.");
            Directory.Delete(Root, true);
        }
    }
}
