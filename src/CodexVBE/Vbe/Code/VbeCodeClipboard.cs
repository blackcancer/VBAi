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
    /// <summary>Instantané du presse-papiers texte avec sa révision et son indicateur de contenu Unicode.</summary>
    internal sealed class CodeClipboardSnapshot
    {
        /// <summary>Révision lue du presse-papiers et empreinte de son texte.</summary>
        /// <value>Valeur utilisée pour refuser une opération fondée sur un instantané périmé.</value>
        public string Version { get; set; }
        /// <summary>Indique si le format Unicode contient du texte.</summary>
        /// <value><see langword="true"/> si le presse-papiers propose le texte Unicode.</value>
        public bool HasText { get; set; }
        /// <summary>Contenu texte capturé lorsqu’il est disponible.</summary>
        /// <value>Texte Unicode ou null si aucun texte n’est présent.</value>
        public string Text { get; set; }
    }
    /// <summary>Frontière de lecture et d’écriture d’un presse-papiers de code.</summary>
    internal interface ICodeClipboard
    {
        /// <summary>Lit un instantané cohérent du presse-papiers.</summary>
        /// <returns>Texte, disponibilité et version du contenu.</returns>
        CodeClipboardSnapshot Read();
        /// <summary>Écrit le texte puis le relit pour vérifier le contenu effectif.</summary>
        /// <param name="text">Texte Unicode à publier.</param>
        /// <returns>Instantané vérifié après écriture.</returns>
        CodeClipboardSnapshot Write(string text);
    }
    /// <summary>Implémente la frontière du presse-papiers WinForms sur le thread STA de VBE.</summary>
    internal sealed class WindowsCodeClipboard : ICodeClipboard
    {
        /// <summary>Lit le numéro de séquence natif du presse-papiers Windows.</summary>
        /// <returns>Numéro incrémenté lors des modifications du presse-papiers.</returns>
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        /// <summary>Lit la révision native du presse-papiers sans remplacer les contrôles de cohérence.</summary>
        internal Func<uint> SequenceNative = GetClipboardSequenceNumber;
        /// <summary>Teste la présence du format Unicode dans le presse-papiers Windows.</summary>
        internal Func<TextDataFormat, bool> ContainsNative = Clipboard.ContainsText;
        /// <summary>Lit le texte au format demandé depuis Windows.</summary>
        internal Func<TextDataFormat, string> GetNative = Clipboard.GetText;
        /// <summary>Écrit le texte au format demandé dans Windows.</summary>
        internal Action<string, TextDataFormat> SetNative = Clipboard.SetText;
        /// <summary>Lit le format Unicode et vérifie que le numéro de séquence n’a pas changé pendant la capture.</summary>
        /// <returns>Instantané contenant la version, la disponibilité et le texte courant.</returns>
        /// <exception cref="InvalidOperationException">L’appel n’est pas sur STA, le presse-papiers a changé ou le texte dépasse un million de caractères.</exception>
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
        /// <summary>Écrit un texte Unicode borné puis confirme sa relecture exacte.</summary>
        /// <param name="text">Texte à copier, entre 1 et 1 048 576 caractères.</param>
        /// <returns>Instantané vérifié après l’écriture.</returns>
        /// <exception cref="ArgumentException">Le texte est vide ou dépasse la limite.</exception>
        /// <exception cref="InvalidOperationException">Le thread n’est pas STA ou la relecture ne correspond pas.</exception>
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
    /// <summary>Copie, coupe et colle une sélection VBA en protégeant l’opération par empreinte et révision clipboard.</summary>
    internal sealed class VbeCodeClipboard
    {
        /// <summary>Transport de lecture et d’écriture des modules VBE.</summary>
        private readonly Func<Request, Response> execute;
        /// <summary>Service qui lit et écrit le contenu texte du presse-papiers.</summary>
        private readonly ICodeClipboard clipboard;
        /// <summary>Crée le service avec le transport VBE et la frontière de presse-papiers à utiliser.</summary>
        /// <param name="execute">Exécuteur des commandes de module.</param>
        /// <param name="clipboard">Implémentation de lecture/écriture du presse-papiers.</param>
        internal VbeCodeClipboard(Func<Request, Response> execute, ICodeClipboard clipboard)
        { this.execute = execute; this.clipboard = clipboard; }
        /// <summary>Lit l’instantané courant du presse-papiers de code.</summary>
        /// <returns>Version, disponibilité du texte et contenu courant.</returns>
        internal object Read() { return clipboard.Read(); }
        /// <summary>Copie, coupe ou colle une plage validée du module selon l’action demandée.</summary>
        /// <param name="request">Projet, module, plage, empreinte source et version clipboard attendue.</param>
        /// <param name="action">copy, cut ou paste.</param>
        /// <returns>Résultat de copie ou état et empreinte de la modification relue.</returns>
        /// <exception cref="ArgumentException">L’action, la plage ou les caractères du presse-papiers sont invalides.</exception>
        /// <exception cref="InvalidOperationException">Le module ou le presse-papiers a changé, ou l’opération hôte a échoué.</exception>
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
        /// <summary>Convertit une plage de caractères à base un avec fin exclusive en décalage de chaîne.</summary>
        /// <param name="code">Code complet utilisant des fins de ligne CRLF.</param>
        /// <param name="request">Positions de début et de fin dans les coordonnées de ligne/colonne.</param>
        /// <param name="start">Reçoit l’offset de départ dans le code.</param>
        /// <param name="length">Reçoit le nombre de caractères sélectionnés.</param>
        /// <exception cref="ArgumentException">La plage dépasse le code ou n’est pas une sélection avant valide.</exception>
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
        /// <summary>Calcule l’empreinte SHA-256 du texte UTF-8 en hexadécimal minuscule.</summary>
        /// <param name="text">Texte du module ou du presse-papiers.</param>
        /// <returns>Empreinte de 64 caractères hexadécimaux.</returns>
        internal static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
    }
}
