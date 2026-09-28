using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class CodeClipboardSnapshot
    {
        public string Version { get; set; }
        public bool HasText { get; set; }
        public string Text { get; set; }
    }
    internal interface ICodeClipboard
    {
        CodeClipboardSnapshot Read();
        CodeClipboardSnapshot Write(string text);
    }
    internal sealed class WindowsCodeClipboard : ICodeClipboard
    {
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        /// <summary>Lit la révision native du presse-papiers sans remplacer les contrôles de cohérence.</summary>
        internal Func<uint> SequenceNative = GetClipboardSequenceNumber;
        /// <summary>Teste la présence du format Unicode dans le presse-papiers Windows.</summary>
        internal Func<TextDataFormat, bool> ContainsNative = Clipboard.ContainsText;
        /// <summary>Lit le texte au format demandé depuis Windows.</summary>
        internal Func<TextDataFormat, string> GetNative = Clipboard.GetText;
        /// <summary>Écrit le texte au format demandé dans Windows.</summary>
        internal Action<string, TextDataFormat> SetNative = Clipboard.SetText;
        public CodeClipboardSnapshot Read()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Clipboard access requires the VBE STA.");
            uint sequence = SequenceNative();
            bool hasText = ContainsNative(TextDataFormat.UnicodeText);
            string text = hasText ? GetNative(TextDataFormat.UnicodeText) : null;
            if (sequence != SequenceNative()) throw new InvalidOperationException("Clipboard changed during inspection; read it again.");
            if (text != null && text.Length > 1024 * 1024) throw new InvalidOperationException("Clipboard text exceeds one million characters.");
            return new CodeClipboardSnapshot { HasText = hasText, Text = text,
                Version = sequence.ToString(CultureInfo.InvariantCulture) + ":" + VbeCodeClipboard.Hash(text ?? "") };
        }
        public CodeClipboardSnapshot Write(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length > 1024 * 1024) throw new ArgumentException("Copy requires 1 to 1048576 characters.");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Clipboard access requires the VBE STA.");
            SetNative(text, TextDataFormat.UnicodeText);
            var result = Read();
            if (!result.HasText || result.Text != text) throw new InvalidOperationException("Clipboard readback differs; source was not cut.");
            return result;
        }
    }
    internal sealed class VbeCodeClipboard
    {
        private readonly Func<Request, Response> execute;
        private readonly ICodeClipboard clipboard;
        internal VbeCodeClipboard(Func<Request, Response> execute, ICodeClipboard clipboard)
        { this.execute = execute; this.clipboard = clipboard; }
        internal object Read() { return clipboard.Read(); }
        internal object Edit(Request request, string action)
        {
            if (action != "copy" && action != "cut" && action != "paste") throw new ArgumentException("Unknown clipboard action.");
            if (action != "copy")
            {
                var state = execute(new Request { Command = "debug_state", Project = request.Project });
                if (!state.Ok || (int)((dynamic)state.Data).Mode != 2) throw new InvalidOperationException("Cut and paste require design mode.");
            }
            var read = execute(new Request { Command = "read_module", Project = request.Project, Module = request.Module });
            if (!read.Ok) throw new InvalidOperationException(read.Error);
            string before = (string)((dynamic)read.Data).Code;
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256) || !string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int start, length;
            Range(before, request, out start, out length);
            if (action != "paste" && length == 0) throw new ArgumentException("Copy and cut require a nonempty range.");
            CodeClipboardSnapshot snapshot;
            string replacement = "";
            if (action == "paste")
            {
                snapshot = clipboard.Read();
                if (string.IsNullOrWhiteSpace(request.ExpectedClipboardVersion) || snapshot.Version != request.ExpectedClipboardVersion)
                    throw new InvalidOperationException("Clipboard changed since inspection; read_code_clipboard again.");
                if (!snapshot.HasText) throw new InvalidOperationException("Clipboard does not contain Unicode text.");
                replacement = snapshot.Text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
                if (replacement.Any(c => char.IsControl(c) && c != '\r' && c != '\n' && c != '\t'))
                    throw new ArgumentException("Clipboard contains unsupported control characters.");
            }
            else snapshot = clipboard.Write(before.Substring(start, length));
            if (action == "copy") return new { Copied = true, Characters = length, ClipboardVersion = snapshot.Version, SourceChanged = false };
            string after = before.Remove(start, length).Insert(start, replacement);
            if (after == before) return new { Applied = false, Verified = true, Sha256 = Hash(before), ClipboardVersion = snapshot.Version, SourceChanged = false };
            var result = execute(new Request { Command = "replace_lines", Project = request.Project, Module = request.Module,
                StartLine = 1, Count = CodeRollback.Lines(before).Length, Text = after, ExpectedSha256 = request.ExpectedSha256 });
            if (!result.Ok) throw new InvalidOperationException(result.Error + (action == "cut" ? " Selected text was copied to the clipboard before the edit failed." : ""));
            var actual = execute(new Request { Command = "read_module", Project = request.Project, Module = request.Module });
            if (!actual.Ok) return new { Applied = true, Verified = false, Error = actual.Error, NextRead = "read_module" };
            string code = (string)((dynamic)actual.Data).Code;
            return new { Applied = true, Verified = code == after, Sha256 = Hash(code), ClipboardVersion = snapshot.Version,
                SourceChanged = code != before, NextRead = code == after ? null : "read_module: VBE normalized or changed the requested text" };
        }
        internal static void Range(string code, Request request, out int start, out int length)
        {
            string[] lines = code.Length == 0 ? new[] { "" } : code.Split(new[] { "\r\n" }, StringSplitOptions.None);
            int sl = request.StartLine, el = request.EndLine, sc = request.StartColumn, ec = request.EndColumn;
            if (sl < 1 || el < sl || el > lines.Length || sc < 1 || ec < 1 ||
                sc > lines[sl - 1].Length + 1 || ec > lines[el - 1].Length + 1 || (sl == el && ec < sc))
                throw new ArgumentException("A valid forward one-based character range with an exclusive end is required.");
            start = lines.Take(sl - 1).Sum(x => x.Length + 2) + sc - 1;
            int end = lines.Take(el - 1).Sum(x => x.Length + 2) + ec - 1;
            length = end - start;
        }
        internal static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
    }
}
