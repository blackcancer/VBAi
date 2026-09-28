using System;
using System.Security.Cryptography;
using System.Text;

namespace CodexVBE
{
    internal interface IEditorModule
    {
        string Name { get; }
        string Key { get; }
        string Read();
        bool CanWrite { get; }
        string Write(string expected, string text);
        void ShowNative(int line, int column);
    }

    /// <summary>A revisioned buffer. External edits never silently replace a dirty buffer.</summary>
    internal sealed class EditorDocument
    {
        internal const int MaxLength = 2 * 1024 * 1024;
        internal readonly string Id = Guid.NewGuid().ToString("N");
        internal readonly IEditorModule Module;
        internal string RecoveryKey { get; private set; }
        internal string Baseline { get; private set; }
        internal string Text { get; private set; }
        internal string Native { get; private set; }
        internal bool Conflict { get; private set; }
        internal bool Dirty => Text != Baseline;
        internal bool Writable => Module.CanWrite;
        internal EditorDocument(IEditorModule module)
        { Module = module; RecoveryKey = module.Key; Baseline = Text = Native = Normalize(module.Read()); Validate(Text); }
        internal void Edit(string value) { Validate(value); Text = Normalize(value); }
        internal void Restore(string baseline, string draft)
        { Validate(baseline); Validate(draft); Baseline = Normalize(baseline); Text = Normalize(draft); Observe(); }
        internal string Observe()
        {
            Native = Normalize(Module.Read()); RecoveryKey = Module.Key;
            Conflict = Dirty && Native != Baseline && Native != Text;
            return !Dirty && Native != Baseline ? Native : null;
        }
        // Call only after the renderer accepted the guarded replacement.
        internal void AcceptRemote(string code) { Baseline = Text = Native = Normalize(code); Conflict = false; }
        internal void Acknowledge(string code, string capturedText)
        { Baseline = Native = Normalize(code); if (Text == capturedText) Text = Baseline; Conflict = false; }
        internal string Synchronize(EditorSyncPlan plan = null)
        {
            Observe();
            if (!Dirty) return Text;
            if (Conflict) throw new InvalidOperationException("The module changed in VBA. Resolve the conflict first.");
            if (!Writable) throw new InvalidOperationException("VBA is running, paused or unavailable. Your draft is preserved.");
            if (Native == Text) { Baseline = Text; return Text; }
            if (plan != null && (plan.Before != Baseline || plan.After != Text))
                throw new InvalidOperationException("The draft changed while synchronization was being prepared.");
            string actual = Normalize(Module is EditorVbeModule nativeModule
                ? nativeModule.WritePrepared(Baseline, Text, plan) : Module.Write(Baseline, Text));
            // Advance the baseline only after a successful write and readback.
            Baseline = Native = actual;
            return actual;
        }
        internal string ResolveWithDraft(string reviewedNative)
        {
            Observe();
            if (Native != reviewedNative) throw new InvalidOperationException("The module changed again. Compare with VBA before resolving.");
            if (!Writable) throw new InvalidOperationException("VBA is running, paused or unavailable. Your draft is preserved.");
            string actual = Normalize(Module.Write(reviewedNative, Text));
            Baseline = Native = actual; Conflict = false; return actual;
        }
        internal static string Normalize(string text) => (text ?? "").Replace("\r\n", "\n").Replace("\r", "\n");
        internal static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
        internal static void Validate(string text)
        { if (text == null || text.Length > MaxLength || text.IndexOf('\0') >= 0) throw new InvalidOperationException("The editor document is invalid or too large."); }
        internal static Tuple<int, int, string> Difference(string before, string after)
        {
            var a = before.Length == 0 ? new string[0] : Normalize(before).Split('\n');
            var b = after.Length == 0 ? new string[0] : Normalize(after).Split('\n');
            int prefix = 0, suffix = 0;
            while (prefix < a.Length && prefix < b.Length && a[prefix] == b[prefix]) prefix++;
            while (suffix < a.Length - prefix && suffix < b.Length - prefix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix]) suffix++;
            return Tuple.Create(prefix + 1, a.Length - prefix - suffix, string.Join("\r\n", b, prefix, b.Length - prefix - suffix));
        }
    }
}
