using System;

namespace CodexVBE
{
    /// <summary>Accesses one live COM component and writes only the changed line range.</summary>
    internal sealed class VbeEditorDocument
    {
        private readonly dynamic vbe;
        private readonly dynamic project;
        private readonly dynamic component;
        internal EditorBuffer Buffer { get; }
        internal string Title => (string)project.Name + " / " + (string)component.Name;
        internal bool CanWrite
        {
            get { try { ValidateIdentity(); return (int)project.Mode == 2; } catch { return false; } }
        }

        internal VbeEditorDocument(object vbe, object component)
        {
            this.vbe = vbe;
            this.component = component;
            project = this.component.Collection.Parent;
            Buffer = new EditorBuffer(Read, Write, () => CanWrite);
        }

        private void ValidateIdentity()
        {
            bool found = false;
            foreach (object candidate in vbe.VBProjects)
                if (ReferenceEquals(candidate, (object)project)) { found = true; break; }
            if (!found || !ReferenceEquals((object)project.VBComponents.Item((string)component.Name), (object)component))
                throw new InvalidOperationException(UiText.Get("The source project or module is no longer available."));
        }

        private string Read()
        {
            ValidateIdentity();
            dynamic module = component.CodeModule;
            int count = module.CountOfLines;
            return count == 0 ? "" : (string)module.Lines[1, count];
        }

        private void Write(string text)
        {
            if (!CanWrite) throw new InvalidOperationException(UiText.Get("Editing requires design mode."));
            dynamic module = component.CodeModule;
            string original = Read();
            string[] before = Lines(original), after = Lines(text);
            int prefix = 0, suffix = 0;
            while (prefix < before.Length && prefix < after.Length && before[prefix] == after[prefix]) prefix++;
            while (suffix < before.Length - prefix && suffix < after.Length - prefix &&
                before[before.Length - 1 - suffix] == after[after.Length - 1 - suffix]) suffix++;
            int remove = before.Length - prefix - suffix;
            int insert = after.Length - prefix - suffix;
            if (remove > 0) module.DeleteLines(prefix + 1, remove);
            if (insert > 0) module.InsertLines(prefix + 1, string.Join("\r\n", after, prefix, insert));
        }

        private static string[] Lines(string text) => string.IsNullOrEmpty(text) ? new string[0] : text.Replace("\r\n", "\n").Split('\n');
    }
}
