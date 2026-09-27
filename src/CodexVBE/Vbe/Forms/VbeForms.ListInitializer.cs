using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // Persist list values through VBA code. Designer.List is intentionally
        // left untouched because live designer items do not survive SaveAs.
        public object SetListInitializer(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) ||
                string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) ||
                string.IsNullOrWhiteSpace(request.ExpectedSha256) || request.Items == null)
                throw new ArgumentException("Project, Form, ControlPath, ExpectedTreeVersion, ExpectedSha256 and Items are required.");
            if (request.Items.Length > 64)
                throw new ArgumentException("At most 64 list items can be initialized.");
            foreach (string item in request.Items)
                if (item == null || item.Length > 256 || item.Any(char.IsControl))
                    throw new ArgumentException("Each item must be a non-null, single-line string of at most 256 characters.");
            Match pathMatch = Regex.Match(request.ControlPath, @"^Controls/([A-Za-z][A-Za-z0-9_]*)$");
            if (!pathMatch.Success)
                throw new ArgumentException("The initial prototype requires a top-level control path.");
            string name = pathMatch.Groups[1].Value;

            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string type = TypeDescriptor.GetClassName(target);
            if (!string.Equals(type, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(type, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ControlPath must identify a ComboBox or ListBox.");
            dynamic list = target;
            if (!string.IsNullOrWhiteSpace(Convert.ToString(list.RowSource, CultureInfo.InvariantCulture)))
                throw new InvalidOperationException("The list is bound to RowSource; a generated initializer is unavailable.");
            if (Convert.ToInt32(list.ColumnCount, CultureInfo.InvariantCulture) != 1)
                throw new InvalidOperationException("The initializer is restricted to one-column lists.");

            dynamic module = form.CodeModule;
            string before = ReadFormCode(module);
            if (!string.Equals(FormCodeSha(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form code changed since it was read.");
            string beginPrefix = "' CodexVBE BEGIN LIST " + request.ControlPath + " SHA256=";
            string end = "' CodexVBE END LIST " + request.ControlPath;
            string[] generated = GenerateListBlock(name, request.Items, beginPrefix, end);
            string begin = generated[0].Trim();
            var allBefore = CodeLines(before);
            int oldBegin = FindMarkerPrefix(allBefore, beginPrefix);
            int oldEnd = FindMarker(allBefore, end);
            if ((oldBegin < 0) != (oldEnd < 0) || (oldBegin >= 0 && oldEnd <= oldBegin))
                throw new InvalidOperationException("The managed list block is incomplete or out of order.");
            if (oldBegin >= 0)
                ValidateManagedBlock(allBefore, oldBegin, oldEnd, name, beginPrefix);

            bool started = false;
            try
            {
                int bodyLine = FindInitializeBody(module);
                if (bodyLine == 0)
                {
                    if (oldBegin >= 0)
                        throw new InvalidOperationException("A managed list block exists outside UserForm_Initialize.");
                    started = true;
                    module.CreateEventProc("Initialize", "UserForm");
                    bodyLine = FindInitializeBody(module);
                    if (bodyLine == 0)
                        throw new InvalidOperationException("CreateEventProc did not create UserForm_Initialize.");
                }
                string afterEventCreation = ReadFormCode(module);
                if (oldBegin >= 0 && !MarkerWithinProcedure(module, oldBegin + 1, oldEnd + 1))
                    throw new InvalidOperationException("The managed list block is outside UserForm_Initialize.");
                int endSubLine = FindEndSubLine(module, bodyLine);
                if (oldBegin >= 0 && oldEnd - oldBegin + 1 == generated.Length &&
                    allBefore.Skip(oldBegin).Take(generated.Length).SequenceEqual(generated))
                    return new { Project = request.Project, Form = request.Form,
                        ControlPath = request.ControlPath, Mechanism = "UserForm_Initialize",
                        Applied = false, Verified = true, VerificationPending = false,
                        UserCodePreserved = true, ItemsWritten = request.Items.Length,
                        Sha256Before = request.ExpectedSha256, Sha256After = request.ExpectedSha256,
                        RuntimeVerificationPending = true, NextRead = "read_module" };
                if (oldBegin < 0)
                {
                    started = true;
                    module.InsertLines(endSubLine, string.Join("\r\n", generated));
                }
                else
                {
                    // Insert before the old block's end, then remove only the
                    // original marked lines. An insertion error leaves them intact.
                    int oldCount = oldEnd - oldBegin + 1;
                    int insertionLine = oldEnd + 2;
                    started = true;
                    module.InsertLines(insertionLine, string.Join("\r\n", generated));
                    module.DeleteLines(oldBegin + 1, oldCount);
                }
                string after = ReadFormCode(module);
                string[] afterLines = CodeLines(after);
                int finalBegin = FindMarker(afterLines, begin);
                int finalEnd = FindMarker(afterLines, end);
                bool exact = finalBegin >= 0 && finalEnd - finalBegin + 1 == generated.Length &&
                    afterLines.Skip(finalBegin).Take(generated.Length).SequenceEqual(generated) &&
                    MarkerWithinProcedure(module, finalBegin + 1, finalEnd + 1);
                bool preserved = oldBegin < 0
                    ? ContainsOrderedLines(afterLines, CodeLines(afterEventCreation))
                    : StripBlock(allBefore, oldBegin, oldEnd)
                        .SequenceEqual(StripBlock(afterLines, finalBegin, finalEnd));
                return new { Project = request.Project, Form = request.Form,
                    ControlPath = request.ControlPath, Mechanism = "UserForm_Initialize",
                    Applied = true, Verified = exact && preserved,
                    VerificationPending = !(exact && preserved),
                    UserCodePreserved = preserved, ItemsWritten = request.Items.Length,
                    Sha256Before = request.ExpectedSha256, Sha256After = FormCodeSha(after),
                    RuntimeVerificationPending = true, NextRead = "read_module" };
            }
            catch (Exception ex)
            {
                if (!started) throw;
                string currentSha = null;
                try { currentSha = FormCodeSha(ReadFormCode(module)); } catch { }
                return new { Project = request.Project, Form = request.Form,
                    ControlPath = request.ControlPath, Mechanism = "UserForm_Initialize",
                    Applied = (bool?)null, Verified = false, VerificationPending = true,
                    NativeError = ex.Message, Sha256Before = request.ExpectedSha256,
                    Sha256After = currentSha, RuntimeVerificationPending = true,
                    NextRead = "read_module" };
            }
        }

        private static string[] GenerateListBlock(string name, string[] items, string beginPrefix, string end)
        {
            var body = new List<string> { "    Me." + name + ".Clear" };
            foreach (string item in items)
                body.Add("    Me." + name + ".AddItem \"" + item.Replace("\"", "\"\"") + "\"");
            var lines = new List<string> { "    " + beginPrefix + FormCodeSha(string.Join("\r\n", body)) };
            lines.AddRange(body);
            lines.Add("    " + end);
            return lines.ToArray();
        }

        private static void ValidateManagedBlock(string[] lines, int begin, int end,
            string name, string beginPrefix)
        {
            string marker = lines[begin].Trim();
            string storedSha = marker.Substring(beginPrefix.Length);
            if (!Regex.IsMatch(storedSha, "^[0-9a-f]{64}$") || end < begin + 2 || end - begin > 66)
                throw new InvalidOperationException("The managed list block has invalid metadata.");
            string[] body = lines.Skip(begin + 1).Take(end - begin - 1).ToArray();
            if (!string.Equals(body[0], "    Me." + name + ".Clear", StringComparison.Ordinal))
                throw new InvalidOperationException("The managed list block has been edited outside the generated form.");
            string itemPattern = "^    Me\\." + Regex.Escape(name) + "\\.AddItem \"(?:[^\"]|\"\")*\"$";
            if (body.Skip(1).Any(line => !Regex.IsMatch(line, itemPattern)) ||
                !string.Equals(FormCodeSha(string.Join("\r\n", body)), storedSha, StringComparison.Ordinal))
                throw new InvalidOperationException("The managed list block has been edited; user code will not be overwritten.");
        }

        private static string ReadFormCode(dynamic module)
        {
            int count = (int)module.CountOfLines;
            return count == 0 ? string.Empty : (string)module.Lines[1, count];
        }

        private static string FormCodeSha(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code)))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static string[] CodeLines(string code)
        {
            if (code.Length == 0) return new string[0];
            string[] lines = Regex.Split(code, "\r\n|\n|\r");
            return lines.Length > 0 && lines[lines.Length - 1].Length == 0
                ? lines.Take(lines.Length - 1).ToArray() : lines;
        }

        private static int FindMarker(string[] lines, string marker)
        {
            int found = -1;
            for (int i = 0; i < lines.Length; i++)
                if (string.Equals(lines[i].Trim(), marker, StringComparison.Ordinal))
                {
                    if (found >= 0) throw new InvalidOperationException("A managed list marker is duplicated: " + marker);
                    found = i;
                }
            return found;
        }

        private static int FindMarkerPrefix(string[] lines, string prefix)
        {
            int found = -1;
            for (int i = 0; i < lines.Length; i++)
                if (lines[i].Trim().StartsWith(prefix, StringComparison.Ordinal))
                {
                    if (found >= 0) throw new InvalidOperationException("A managed list marker is duplicated: " + prefix);
                    found = i;
                }
            return found;
        }

        private static int FindInitializeBody(dynamic module)
        {
            try { return (int)module.ProcBodyLine["UserForm_Initialize", 0]; }
            catch (System.Runtime.InteropServices.COMException) { return 0; }
        }

        private static int FindEndSubLine(dynamic module, int bodyLine)
        {
            int start = (int)module.ProcStartLine["UserForm_Initialize", 0];
            int count = (int)module.ProcCountLines["UserForm_Initialize", 0];
            for (int line = bodyLine + 1; line < start + count; line++)
                if (string.Equals(((string)module.Lines[line, 1]).Trim(), "End Sub",
                    StringComparison.OrdinalIgnoreCase)) return line;
            throw new InvalidOperationException("UserForm_Initialize has no unambiguous End Sub line.");
        }

        private static bool MarkerWithinProcedure(dynamic module, int beginLine, int endLine)
        {
            int start = (int)module.ProcStartLine["UserForm_Initialize", 0];
            int count = (int)module.ProcCountLines["UserForm_Initialize", 0];
            return beginLine > start && endLine < start + count;
        }

        private static bool ContainsOrderedLines(string[] actual, string[] expected)
        {
            int match = 0;
            foreach (string line in actual)
                if (match < expected.Length && line == expected[match]) match++;
            return match == expected.Length;
        }

        private static IEnumerable<string> StripBlock(string[] lines, int begin, int end)
        {
            if (begin < 0 || end < begin) return lines;
            return lines.Where((line, index) => index < begin || index > end);
        }
    }
}
