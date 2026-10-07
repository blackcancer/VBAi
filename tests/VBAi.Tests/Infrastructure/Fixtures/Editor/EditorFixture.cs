using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
namespace VBAi.Tests.Infrastructure
{
    internal sealed class EditorFixture : IEditorModule, IDisposable
    {
        private readonly SynchronizationContext previousContext = SynchronizationContext.Current;
        internal EditorFixture()
        {
            // Real WebView tests need a live WinForms dispatcher even after unrelated
            // Designer/WPF scenarios have run on the test runner's STA thread.
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        }
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "vbai-editor-" + Guid.NewGuid().ToString("N"));
        internal string Code = "Option Explicit\nPublic Sub Hello()\n    Debug.Print 1\nEnd Sub";
        internal int Writes;
        internal bool Fail;
        internal string DisplayName = "Fixture · Module1";
        public string Name => DisplayName;
        internal bool Closed; public string Key => Closed ? throw new InvalidOperationException("Closed") : Root;
        public bool CanWrite { get; set; } = true;
        public string Read() => Code;
        public string Write(string expected, string text)
        { if (Code != expected || !CanWrite || Fail) throw new InvalidOperationException("Write refused"); Writes++; return Code = text; }
        public void ShowNative(int line, int column) { }
        public void Dispose()
        { try { if (Directory.Exists(Root)) Directory.Delete(Root, true); } finally { SynchronizationContext.SetSynchronizationContext(previousContext); } }
    }
}
