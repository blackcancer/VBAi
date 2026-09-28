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
        /// <summary>Manifest resource name of the native renderer payload.</summary>
private const string ResourceName = "CodexVBE.Native.Renderer.dll";
        /// <summary>Native renderer entry point that accepts one window handle and returns a status code.</summary>
        /// <param name="window">Window handle passed to the native entry point.</param><returns>Native status code.</returns>
[UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint WindowCall(IntPtr window);
        /// <summary>Native renderer entry point with no arguments and a status result.</summary><returns>Native status code.</returns>
[UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint SimpleCall();
        /// <summary>Native renderer entry point that fills the ABI-compatible status structure.</summary>
        /// <param name="status">Structure receiving renderer state and counters.</param><returns>Native status code.</returns>
[UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate uint StatusCall(ref RendererStatus status);
        /// <summary>Native ABI structure containing compatibility data and rendering diagnostics.</summary>
[StructLayout(LayoutKind.Sequential)] internal struct RendererStatus
        {
            /// <summary>Structure size, ABI version, active state, registered windows, and import totals.</summary>
internal uint Size, Abi, Active, Windows, Imports, RestoredImports;
            /// <summary>Pattern, image, text, fill, unsupported-item, and failure counters.</summary>
internal uint Patterns, Images, Text, Fills, Unsupported, Failures;
        }
        /// <summary>Loads the native renderer library from a verified cache path.</summary>
        /// <param name="path">Absolute module path.</param><param name="file">Reserved file handle.</param><param name="flags">Load behavior flags.</param>
        /// <returns>Module handle, or zero on failure.</returns>
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
        /// <summary>Resolves an exported function from a loaded native module.</summary>
        /// <param name="module">Loaded module handle.</param><param name="name">Export name.</param>
        /// <returns>Function pointer, or zero on failure.</returns>
[DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr module, string name);
        /// <summary>Unloads a native module after its hooks and callbacks are no longer active.</summary>
        /// <param name="module">Native module handle.</param><returns>Whether the module was released.</returns>
[DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr module);
        /// <summary>Loaded renderer module retained while hooks may reference native code.</summary>
private static IntPtr module;
        /// <summary>Resolved native renderer start and window-registration entry points.</summary>
private static WindowCall start, register;
        /// <summary>Resolved native renderer stop and import-refresh entry points.</summary>
private static SimpleCall stop, refresh;
        /// <summary>Resolved native status-query entry point.</summary>
private static StatusCall query;
        /// <summary>Cached state indicating whether native hooks are active.</summary>
private static bool active;
        /// <summary>Gets whether the native renderer currently reports an active hook.</summary>
        /// <value><see langword="true"/> while the renderer's start/stop state is active.</value>
internal static bool Active => active;
        /// <summary>Stores the supports loader host used by VbeNativeRenderer.</summary>
internal static Func<bool> SupportsLoaderHost = () => Environment.Is64BitProcess;
        /// <summary>Stores the open payload used by VbeNativeRenderer.</summary>
internal static Func<Stream> OpenPayload = () => typeof(VbeNativeRenderer).Assembly.GetManifestResourceStream(ResourceName);
        /// <summary>Stores the cache root used by VbeNativeRenderer.</summary>
internal static string CacheRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "native-renderer");
        /// <summary>Stores the move payload used by VbeNativeRenderer.</summary>
internal static Action<string, string> MovePayload = File.Move;
        /// <summary>Stores the load module used by VbeNativeRenderer.</summary>
internal static Func<string, IntPtr, uint, IntPtr> LoadModule = LoadLibraryEx;
        /// <summary>Stores the find export used by VbeNativeRenderer.</summary>
internal static Func<IntPtr, string, IntPtr> FindExport = GetProcAddress;
        /// <summary>Stores the release module used by VbeNativeRenderer.</summary>
internal static Func<IntPtr, bool> ReleaseModule = FreeLibrary;

        /// <summary>Loads the embedded renderer and starts its hooks for the VBE editor window.</summary>
        /// <param name="editor">Handle of the VBE editor window passed to the native start entry point.</param>
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
        /// <summary>Registers a newly discovered VBE window with the active native renderer.</summary>
        /// <param name="window">Handle of the window to register.</param>
internal static void Register(IntPtr window)
        {
            if (!active) return;
            uint error = register(window);
            if (error == 0 || error == 50) return;
            LoadLog.Write("Native toolbar registration failed: " + error);
            Stop();
        }
        /// <summary>Refreshes the native renderer's import hooks while it is active.</summary>
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
        /// <summary>Reads a diagnostic summary from the native renderer, when its query entry point is available.</summary>
        /// <returns>Human-readable hook and rendering counters, or a query status.</returns>
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
        /// <summary>Extracts, verifies, and loads the embedded renderer and resolves its ABI entry points.</summary>
        /// <exception cref="PlatformNotSupportedException">The current process is not x64.</exception>
        /// <exception cref="FileNotFoundException">The renderer resource is absent from the add-in assembly.</exception>
        /// <exception cref="InvalidDataException">The cached payload or native ABI is invalid.</exception>
private static void EnsureLoaded()
        {
            if (module != IntPtr.Zero) return;
            if (!SupportsLoaderHost()) throw new PlatformNotSupportedException("The native VBE renderer requires an x64 host.");
            byte[] bytes;
            using (Stream resource = OpenPayload())
            {
                if (resource == null) throw new FileNotFoundException("The embedded native renderer is missing. Rebuild the complete add-in.");
                using (var buffer = new MemoryStream()) { resource.CopyTo(buffer); bytes = buffer.ToArray(); }
            }
            string hash = Hash(bytes);
            string directory = Path.Combine(CacheRoot, hash);
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "CodexVBE.Native.dll");
            if (!File.Exists(path))
            {
                string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    File.WriteAllBytes(temporary, bytes);
                    try { MovePayload(temporary, path); }
                    catch (IOException) { if (!File.Exists(path)) throw; }
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            if (!string.Equals(Hash(File.ReadAllBytes(path)), hash, StringComparison.Ordinal))
                throw new InvalidDataException("The cached native renderer differs from the embedded payload.");
            IntPtr loaded = LoadModule(path, IntPtr.Zero, 0x00000100 | 0x00001000);
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
                ReleaseModule(loaded);
                throw;
            }
        }
        /// <summary>Resolves a named native export and marshals it to the requested delegate type.</summary>
        /// <typeparam name="T">Managed delegate type matching the native export signature.</typeparam>
        /// <param name="library">Loaded native module handle.</param>
        /// <param name="name">Export name.</param>
        /// <returns>The marshaled delegate.</returns>
        /// <exception cref="EntryPointNotFoundException">The export is not present in the module.</exception>
private static T Resolve<T>(IntPtr library, string name) where T : class
        {
            IntPtr address = FindExport(library, name);
            if (address == IntPtr.Zero) throw new EntryPointNotFoundException(name);
            return (T)(object)Marshal.GetDelegateForFunctionPointer(address, typeof(T));
        }
        /// <summary>Computes the uppercase hexadecimal SHA-256 digest of a byte array.</summary>
        /// <param name="bytes">Payload bytes to hash.</param>
        /// <returns>Digest without separators.</returns>
private static string Hash(byte[] bytes)
        {
            using (var algorithm = SHA256.Create()) return BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", string.Empty);
        }
    }
}
