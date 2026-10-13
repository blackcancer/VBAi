using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{
    /// <summary>Owns one verified Microsoft Office VBA SIP module reference for a file verification.</summary>
    internal sealed class OfficeVbaSignatureProvider
    {
        /// <summary>Dedicated VBA SIP registered by current Microsoft Office installations.</summary>
        internal static readonly Guid Subject = new Guid("9FA65764-C36F-4319-9737-658A34585BB7");

        /// <summary>Required native operations; a document-only provider cannot satisfy the VBA contract.</summary>
        private static readonly string[] Sections = { "CryptSIPDllVerifyIndirectData", "CryptSIPDllGetSignedDataMsg", "CryptSIPDllCreateIndirectData" };
        /// <summary>Exact VBA exports corresponding to the registered operations.</summary>
        private static readonly string[] Functions = { "MsoVBADigSigVerifyIndirectData", "MsoVBADigSigGetSignedDataMsg", "MsoVBADigSigCreateIndirectData" };
        /// <summary>Search only the target DLL directory and the Windows default DLL directories.</summary>
        internal const uint LoadFlags = 0x1100;

        /// <summary>Reads the registered DLL and function for this exact VBA subject from HKLM x64.</summary>
        internal Func<string, string, string> ReadRegistration = ReadNativeRegistration;
        /// <summary>Reads the existing module path without acquiring or releasing its owner's handle.</summary>
        internal Func<string> ReadLoadedPath = ReadNativeLoadedPath;
        /// <summary>Reads the Click-to-Run installation root when the module is not already loaded.</summary>
        internal Func<string> ReadInstallationRoot = ReadNativeInstallationRoot;
        /// <summary>Requires actual Windows Authenticode trust and the verified Microsoft signer before loading.</summary>
        internal Func<string, bool> VerifyMicrosoftModule = path => VbeSignatureVerifier.VerifyProviderModule(path);
        /// <summary>Acquires exactly one module reference using only the absolute file and restricted search flags.</summary>
        internal Func<string, uint, IntPtr> Load = (path, flags) => LoadLibraryExW(path, IntPtr.Zero, flags);
        /// <summary>Reads the acquired module's actual path so a different loaded module cannot satisfy discovery.</summary>
        internal Func<IntPtr, string> ReadAcquiredPath = ReadNativeModulePath;
        /// <summary>Checks the registered functions against the acquired module without invoking them.</summary>
        internal Func<IntPtr, string, IntPtr> Export = GetProcAddress;
        /// <summary>Releases only the reference acquired by Load, at most once.</summary>
        internal Func<IntPtr, bool> Free = FreeLibrary;

        /// <summary>Acquires the current VBA provider, or returns null when registration, path or trust is unavailable.</summary>
        /// <returns>An owner-thread lease keeping the verified file and module alive through WinTrust state closure.</returns>
        /// <remarks>No registration, trusted-publisher entry or persistent DLL search path is changed.</remarks>
        internal Lease TryAcquire()
        {
            var registrations = new string[Sections.Length];
            for (int index = 0; index < Sections.Length; index++)
            {
                string dll = ReadRegistration(Sections[index], "Dll");
                if (string.IsNullOrWhiteSpace(dll) ||
                    (!IsAbsolutePath(dll) && !dll.Equals("mso.dll", StringComparison.OrdinalIgnoreCase)) ||
                    !Path.GetFileName(dll).Equals("mso.dll", StringComparison.OrdinalIgnoreCase) ||
                    ReadRegistration(Sections[index], "FuncName") != Functions[index]) return null;
                registrations[index] = dll;
            }
            string path = ReadLoadedPath();
            if (string.IsNullOrWhiteSpace(path))
            {
                path = Array.Find(registrations, IsAbsolutePath);
                if (path == null)
                {
                    string root = ReadInstallationRoot();
                    if (string.IsNullOrWhiteSpace(root) || !IsAbsolutePath(root)) return null;
                    path = Path.Combine(root, @"root\vfs\ProgramFilesCommonX64\Microsoft Shared\Office16\MSO.DLL");
                }
            }
            if (!IsAbsolutePath(path) || !Path.GetFileName(path).Equals("mso.dll", StringComparison.OrdinalIgnoreCase) || !File.Exists(path)) return null;
            path = Path.GetFullPath(path);
            foreach (string registration in registrations)
                if (IsAbsolutePath(registration) && !path.Equals(Path.GetFullPath(registration), StringComparison.OrdinalIgnoreCase)) return null;

            FileStream file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            IntPtr module = IntPtr.Zero;
            try
            {
                if (!VerifyMicrosoftModule(path)) return null;
                module = Load(path, LoadFlags);
                if (module == IntPtr.Zero) return null;
                string acquiredPath = ReadAcquiredPath(module);
                if (string.IsNullOrWhiteSpace(acquiredPath) || !IsAbsolutePath(acquiredPath) ||
                    !path.Equals(Path.GetFullPath(acquiredPath), StringComparison.OrdinalIgnoreCase)) return null;
                foreach (string function in Functions)
                    if (Export(module, function) == IntPtr.Zero) return null;
                var lease = new Lease(path, module, file, Free);
                module = IntPtr.Zero;
                file = null;
                return lease;
            }
            finally
            {
                try
                {
                    if (module != IntPtr.Zero && !Free(module))
                        throw new InvalidOperationException("The acquired Office VBA provider module could not be released; no retry was attempted.");
                }
                finally { file?.Dispose(); }
            }
        }

        /// <summary>One acquired provider reference; never releases a borrowed handle or repeats an uncertain release.</summary>
        internal sealed class Lease : IDisposable
        {
            /// <summary>Thread that acquired this module reference and must release it.</summary>
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            /// <summary>One-shot release boundary for the acquisition, never a borrowed module handle.</summary>
            private readonly Func<IntPtr, bool> free;
            /// <summary>Acquired module reference, cleared before release to prevent an uncertain retry.</summary>
            private IntPtr module;
            /// <summary>File lock protecting the verified module bytes until native state closure.</summary>
            private FileStream file;
            /// <summary>Exact verified Microsoft module path held open until verification and provider-state closure finish.</summary>
            internal string Path { get; }
            /// <summary>Takes ownership of exactly one module acquisition and its read-only file lock.</summary>
            internal Lease(string path, IntPtr module, FileStream file, Func<IntPtr, bool> free)
            {
                Path = path;
                this.module = module;
                this.file = file;
                this.free = free;
            }
            /// <summary>Releases this acquisition once on its owner thread and closes the read-only file lock.</summary>
            public void Dispose()
            {
                if (Thread.CurrentThread.ManagedThreadId != owner) throw new InvalidOperationException("The Office VBA provider must be released on its owning thread.");
                IntPtr acquired = module;
                module = IntPtr.Zero;
                try
                {
                    if (acquired != IntPtr.Zero && !free(acquired))
                        throw new InvalidOperationException("Office VBA provider release failed; no retry was attempted.");
                }
                finally { file?.Dispose(); file = null; }
            }
        }

        /// <summary>Rejects drive-relative and current-drive-rooted Windows paths before DLL loading.</summary>
        private static bool IsAbsolutePath(string path)
            => !string.IsNullOrWhiteSpace(path) && (path.StartsWith(@"\\", StringComparison.Ordinal) ||
                (path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' && (path[2] == '\\' || path[2] == '/')));

        /// <summary>Reads one value for the dedicated VBA subject without modifying registration.</summary>
        private static string ReadNativeRegistration(string section, string name)
        {
            using (var registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (var key = registry.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\" + section + "\\" + Subject.ToString("B")))
                return key?.GetValue(name) as string;
        }
        /// <summary>Reads the existing Office installation location without creating registry keys.</summary>
        private static string ReadNativeInstallationRoot()
        {
            using (var registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (var key = registry.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                return key?.GetValue("InstallationPath") as string;
        }
        /// <summary>Observes the loaded Office module's path without acquiring a reference.</summary>
        private static string ReadNativeLoadedPath()
        {
            IntPtr module = GetModuleHandleW("mso.dll");
            return module == IntPtr.Zero ? null : ReadNativeModulePath(module);
        }
        /// <summary>Reads one module filename without changing its reference count or any loader search path.</summary>
        private static string ReadNativeModulePath(IntPtr module)
        {
            var path = new StringBuilder(32768);
            uint count = GetModuleFileNameW(module, path, path.Capacity);
            return count == 0 || count >= path.Capacity ? null : path.ToString();
        }
        /// <summary>Acquires a module using the selected absolute path and bounded search directories.</summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryExW(string path, IntPtr file, uint flags);
        /// <summary>Observes an export address without executing it.</summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr module, string name);
        /// <summary>Releases one acquired reference and reports failure without replay.</summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool FreeLibrary(IntPtr module);
        /// <summary>Obtains a borrowed handle whose reference count is unchanged.</summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandleW(string name);
        /// <summary>Reads the exact module path from a borrowed or acquired handle.</summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern uint GetModuleFileNameW(IntPtr module, StringBuilder path, int capacity);
    }
}
