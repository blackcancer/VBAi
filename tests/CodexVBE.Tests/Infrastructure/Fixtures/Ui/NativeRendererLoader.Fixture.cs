using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal sealed class NativeRendererLoaderFixture : IDisposable
    {
        private const BindingFlags Flags = BindingFlags.Static | BindingFlags.NonPublic;
        private readonly Dictionary<FieldInfo, object> fields = new Dictionary<FieldInfo, object>();
        private readonly Func<bool> host = VbeNativeRenderer.SupportsLoaderHost;
        private readonly Func<Stream> payload = VbeNativeRenderer.OpenPayload;
        private readonly string cache = VbeNativeRenderer.CacheRoot;
        private readonly Action<string, string> move = VbeNativeRenderer.MovePayload;
        private readonly Func<string, IntPtr, uint, IntPtr> load = VbeNativeRenderer.LoadModule;
        private readonly Func<IntPtr, string, IntPtr> find = VbeNativeRenderer.FindExport;
        private readonly Func<IntPtr, bool> release = VbeNativeRenderer.ReleaseModule;
        internal readonly string DirectoryPath = Path.Combine(Path.GetTempPath(), "CodexVBE-loader-" + Guid.NewGuid().ToString("N"));
        internal readonly byte[] Bytes;
        internal readonly string Hash;
        internal readonly List<IntPtr> Released = new List<IntPtr>();
        internal readonly List<IntPtr> Loaded = new List<IntPtr>();
        internal int ResourceReads, Loads, Moves;
        internal string MissingExport;
        internal bool OverrideStatus;
        internal uint StatusError, StatusAbi = 1;
        private readonly Delegate statusExport;
        internal string CachedPayload => Path.Combine(DirectoryPath, Hash, "CodexVBE.Native.dll");

        internal NativeRendererLoaderFixture()
        {
            foreach (string name in new[] { "module", "start", "register", "stop", "refresh", "query", "active" })
            {
                var field = typeof(VbeNativeRenderer).GetField(name, Flags); fields[field] = field.GetValue(null);
                field.SetValue(null, field.FieldType.IsValueType ? Activator.CreateInstance(field.FieldType) : null);
            }
            using (var stream = payload())
            using (var buffer = new MemoryStream()) { stream.CopyTo(buffer); Bytes = buffer.ToArray(); }
            using (var algorithm = SHA256.Create()) Hash = BitConverter.ToString(algorithm.ComputeHash(Bytes)).Replace("-", "");
            var statusType = typeof(VbeNativeRenderer).GetField("query", Flags).FieldType;
            statusExport = Delegate.CreateDelegate(statusType, this, GetType().GetMethod("Status", BindingFlags.Instance | BindingFlags.NonPublic));
            VbeNativeRenderer.CacheRoot = DirectoryPath;
            VbeNativeRenderer.OpenPayload = () => { ResourceReads++; return payload(); };
            VbeNativeRenderer.MovePayload = (from, to) => { Moves++; move(from, to); };
            VbeNativeRenderer.LoadModule = (path, file, flags) =>
            {
                Loads++; Assert.AreEqual(CachedPayload, path); Assert.AreEqual(IntPtr.Zero, file); Assert.AreEqual((uint)0x1100, flags);
                IntPtr module = load(path, file, flags); if (module != IntPtr.Zero) Loaded.Add(module); return module;
            };
            VbeNativeRenderer.FindExport = (module, name) => name == MissingExport ? IntPtr.Zero :
                name == "CodexVbeThemeStatus" && OverrideStatus ? Marshal.GetFunctionPointerForDelegate(statusExport) : find(module, name);
            VbeNativeRenderer.ReleaseModule = module => { Released.Add(module); Loaded.Remove(module); return release(module); };
        }
        private uint Status(ref VbeNativeRenderer.RendererStatus status)
        { status.Abi = StatusAbi; status.Active = 0; return StatusError; }
        internal void EnsureLoaded()
        {
            try { typeof(VbeNativeRenderer).GetMethod("EnsureLoaded", Flags).Invoke(null, null); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        internal object Read(string name) => typeof(VbeNativeRenderer).GetField(name, Flags).GetValue(null);
        internal string Describe() => (string)typeof(VbeNativeRenderer).GetMethod("Describe", Flags).Invoke(null, null);
        internal void Seed(byte[] bytes) { Directory.CreateDirectory(Path.GetDirectoryName(CachedPayload)); File.WriteAllBytes(CachedPayload, bytes); }
        internal void AssertNoTemporaryFiles() { if (Directory.Exists(DirectoryPath)) Assert.AreEqual(0, Directory.GetFiles(DirectoryPath, "*.tmp", SearchOption.AllDirectories).Length); }
        public void Dispose()
        {
            foreach (IntPtr module in Loaded) release(module);
            foreach (var pair in fields) pair.Key.SetValue(null, pair.Value);
            VbeNativeRenderer.SupportsLoaderHost = host; VbeNativeRenderer.OpenPayload = payload; VbeNativeRenderer.CacheRoot = cache;
            VbeNativeRenderer.MovePayload = move; VbeNativeRenderer.LoadModule = load; VbeNativeRenderer.FindExport = find; VbeNativeRenderer.ReleaseModule = release;
            if (Directory.Exists(DirectoryPath)) Directory.Delete(DirectoryPath, true);
            GC.KeepAlive(statusExport);
        }
    }
}
