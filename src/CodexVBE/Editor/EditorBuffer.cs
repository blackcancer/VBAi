using System;

namespace CodexVBE
{
    /// <summary>Maintains a draft and refuses to overwrite a concurrently edited module.</summary>
    internal sealed class EditorBuffer
    {
        private readonly Func<string> read;
        private readonly Action<string> write;
        private readonly Func<bool> canWrite;
        internal string Baseline { get; private set; }
        internal string Draft { get; set; }
        internal bool Dirty => Normalize(Draft) != Normalize(Baseline);

        internal EditorBuffer(Func<string> read, Action<string> write, Func<bool> canWrite)
        {
            this.read = read;
            this.write = write;
            this.canWrite = canWrite;
            Reload();
        }

        internal void Reload() { Baseline = read(); Draft = Baseline; }

        internal void Apply()
        {
            if (!canWrite()) throw new InvalidOperationException(UiText.Get("Editing requires design mode."));
            if (Normalize(read()) != Normalize(Baseline))
                throw new InvalidOperationException(UiText.Get("The module changed in the VBE. Your draft is preserved; copy it before reloading."));
            if (!Dirty) return;
            // Keep both strings intact on failure, including a partial COM write.
            write(Normalize(Draft));
            string actual = read();
            if (Normalize(actual) != Normalize(Draft))
                throw new InvalidOperationException(UiText.Get("The VBE returned different code. Your draft is preserved; review the native module before reloading."));
            Baseline = actual;
            Draft = actual;
        }

        internal static string Normalize(string text) => (text ?? "").Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n').Replace("\n", "\r\n");
    }
}
