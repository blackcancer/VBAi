using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace CodexVBE
{
    /// <summary>Loads the embedded x64 renderer and manages its VBE-thread lifecycle.</summary>
    internal static class VbeNativeRenderer
    {
        private const string ResourceName = "CodexVBE.Native.Renderer.dll";
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint WindowCall(IntPtr window);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint SimpleCall();
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint StatusCall(ref RendererStatus status);
        [StructLayout(LayoutKind.Sequential)] internal struct RendererStatus
        {
            internal uint Size, Abi, Active, Windows, Imports, RestoredImports;
            internal uint Patterns, Images, Text, Fills, Unsupported, Failures;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr module, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr module);
        private static IntPtr module;
        private static WindowCall start, register;
        private static SimpleCall stop, refresh;
        private static StatusCall query;
        private static bool active;
        internal static bool Active => active;

        internal static void Start(IntPtr editor)
        {
            try
            {
                EnsureLoaded();
                uint error = start(editor);
                if (error != 0) throw new Win32Exception((int)error);
                active = true;
                LoadLog.Write("Native toolbar renderer started: " + Describe());
            }
            catch (Exception exception)
            {
                Stop();
                LoadLog.Write("Native toolbar renderer unavailable; legacy recovery retained: " + exception.Message);
            }
        }
        internal static void Register(IntPtr window)
        {
            if (!active) return;
            uint error = register(window);
            if (error == 0 || error == 50) return;
            LoadLog.Write("Native toolbar registration failed: " + error);
            Stop();
        }
        internal static void Refresh()
        {
            if (!active) return;
            uint error = refresh();
            if (error == 0) return;
            LoadLog.Write("Native toolbar import refresh failed: " + error);
            Stop();
        }
        /// <summary>Confirms native hook removal before callers release their managed guards.</summary>
        /// <returns>True only for a confirmed stop or an unloaded, inactive renderer; false preserves state and resources.</returns>
        internal static bool Stop()
        {
            if (stop == null)
            {
                if (!active) return true;
                LoadLog.Write("Native toolbar renderer stop unavailable; active state and managed guards must be retained.");
                return false;
            }
            uint error;
            try { error = stop(); }
            catch (Exception exception)
            {
                LoadLog.Write("Native toolbar renderer stop could not be confirmed: " + exception.GetType().Name +
                    "; active state and managed guards must be retained. " + exception.Message);
                return false;
            }
            if (error != 0)
            {
                string reason = error == 1444 ? "ERROR_INVALID_THREAD_ID; stop must run on the owning VBE UI thread" :
                    "native hook restoration was not confirmed";
                LoadLog.Write("Native toolbar renderer stop failed: status=" + error + "; " + reason +
                    "; active state and managed guards must be retained.");
                return false;
            }
            active = false;
            LoadLog.Write("Native toolbar renderer stopped: status=0; " + Describe());
            return true;
        }
        private static string Describe()
        {
            if (query == null) return "not loaded";
            var status = new RendererStatus { Size = (uint)Marshal.SizeOf(typeof(RendererStatus)) };
            uint error = query(ref status);
            if (error != 0) return "query status=" + error;
            return "active=" + status.Active + ", windows=" + status.Windows + ", imports=" + status.Imports +
                ", restored=" + status.RestoredImports + ", patterns=" + status.Patterns + ", icons=" + status.Images +
                ", text=" + status.Text + ", fills=" + status.Fills + ", unsupported=" + status.Unsupported + ", failures=" + status.Failures;
        }
        private static void EnsureLoaded()
        {
            if (module != IntPtr.Zero) return;
            if (!Environment.Is64BitProcess) throw new PlatformNotSupportedException("The native VBE renderer requires an x64 host.");
            byte[] bytes;
            using (Stream resource = typeof(VbeNativeRenderer).Assembly.GetManifestResourceStream(ResourceName))
            {
                if (resource == null) throw new FileNotFoundException("The embedded native renderer is missing. Rebuild the complete add-in.");
                using (var buffer = new MemoryStream()) { resource.CopyTo(buffer); bytes = buffer.ToArray(); }
            }
            string hash = Hash(bytes);
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "native-renderer", hash);
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "CodexVBE.Native.dll");
            if (!File.Exists(path))
            {
                string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    File.WriteAllBytes(temporary, bytes);
                    try { File.Move(temporary, path); }
                    catch (IOException) { if (!File.Exists(path)) throw; }
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            if (!string.Equals(Hash(File.ReadAllBytes(path)), hash, StringComparison.Ordinal))
                throw new InvalidDataException("The cached native renderer differs from the embedded payload.");
            IntPtr loaded = LoadLibraryEx(path, IntPtr.Zero, 0x00000100 | 0x00001000);
            if (loaded == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                var loadedStart = Resolve<WindowCall>(loaded, "CodexVbeThemeStart");
                var loadedRegister = Resolve<WindowCall>(loaded, "CodexVbeThemeRegister");
                var loadedStop = Resolve<SimpleCall>(loaded, "CodexVbeThemeStop");
                var loadedRefresh = Resolve<SimpleCall>(loaded, "CodexVbeThemeRefresh");
                var loadedQuery = Resolve<StatusCall>(loaded, "CodexVbeThemeStatus");
                var status = new RendererStatus { Size = (uint)Marshal.SizeOf(typeof(RendererStatus)) };
                uint statusError = loadedQuery(ref status);
                if (statusError != 0 || status.Abi != 1)
                    throw new InvalidDataException("The native renderer ABI does not match this add-in.");
                start = loadedStart; register = loadedRegister; stop = loadedStop;
                refresh = loadedRefresh; query = loadedQuery;
                // Retain after validation. Start pins the module because another
                // thread may hold a previously fetched IAT function pointer.
                module = loaded;
            }
            catch
            {
                // Start has not run: no hooks or subclasses can reference this module.
                FreeLibrary(loaded);
                throw;
            }
        }
        private static T Resolve<T>(IntPtr library, string name) where T : class
        {
            IntPtr address = GetProcAddress(library, name);
            if (address == IntPtr.Zero) throw new EntryPointNotFoundException(name);
            return (T)(object)Marshal.GetDelegateForFunctionPointer(address, typeof(T));
        }
        private static string Hash(byte[] bytes)
        {
            using (var algorithm = SHA256.Create()) return BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", string.Empty);
        }
    }
}
