using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed partial class VbeCodeNavigation
    {
        private readonly dynamic vbe;
        private readonly VbeForms forms;

        public VbeCodeNavigation(object vbe, VbeForms forms) { this.vbe = vbe; this.forms = forms; }

        public object CreateEventProcedure(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.EventName) ||
                !Regex.IsMatch(request.EventName, @"^[A-Za-z][A-Za-z0-9_]*$") ||
                string.IsNullOrWhiteSpace(request.ObjectName) ||
                !Regex.IsMatch(request.ObjectName, @"^[A-Za-z_][A-Za-z0-9_]*$") ||
                string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("EventName, ObjectName, ExpectedSha256 and ExpectedTreeVersion are required.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            dynamic component = null;
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, request.Form, StringComparison.OrdinalIgnoreCase))
                { component = candidate; break; }
            if (component == null || (int)component.Type != 3)
                throw new InvalidOperationException("The target must be an existing UserForm.");
            dynamic tree = forms.Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!string.Equals(request.ObjectName, "UserForm", StringComparison.OrdinalIgnoreCase) &&
                CountControls((IEnumerable)tree.Controls, request.ObjectName) != 1)
                throw new InvalidOperationException("ObjectName must identify exactly one control in form_tree or UserForm.");
            dynamic module = component.CodeModule;
            string before = Code(module, (int)module.CountOfLines);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form code changed since it was read.");
            string procedureName = request.ObjectName + "_" + request.EventName;
            if (Regex.IsMatch(before, @"(?im)^\s*(?:(?:Private|Public|Friend|Static)\s+)?Sub\s+" +
                Regex.Escape(procedureName) + @"\b"))
                throw new InvalidOperationException("The event procedure already exists: " + procedureName);
            int bodyLine = (int)module.CreateEventProc(request.EventName, request.ObjectName);
            string after = Code(module, (int)module.CountOfLines);
            int kind = 0;
            string actual = (string)module.ProcOfLine[bodyLine, ref kind];
            if (bodyLine < 1 || string.Equals(after, before, StringComparison.Ordinal) || kind != 0 ||
                !string.Equals(actual, procedureName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The created event procedure could not be verified in the code module.");
            return new { Project = request.Project, Form = request.Form,
                request.ObjectName, request.EventName, Procedure = actual,
                BodyLine = bodyLine, Sha256 = Hash(after), Code = after };
        }

        public object CreateProcedure(Request request)
        {
            string text = ValidateProcedureText(request);
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            dynamic module = GetModule(project, request.Module);
            int componentType = (int)module.Parent.Type;
            if (componentType != 1 && componentType != 2)
                throw new InvalidOperationException("create_procedure targets a standard or class module.");
            int countBefore = (int)module.CountOfLines;
            string before = Code(module, countBefore);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int existingBody = 0;
            try { existingBody = (int)module.ProcBodyLine[request.Procedure, request.ProcKind]; }
            catch { }
            if (existingBody > 0)
                throw new InvalidOperationException("The procedure already exists.");
            string insertion = (countBefore > 0 ? "\r\n" : "") + text.Replace("\n", "\r\n");
            try
            {
                module.InsertLines(countBefore + 1, insertion);
                int countAfter = (int)module.CountOfLines;
                int bodyLine = (int)module.ProcBodyLine[request.Procedure, request.ProcKind];
                int actualKind = request.ProcKind;
                string actual = (string)module.ProcOfLine[bodyLine, ref actualKind];
                if (bodyLine <= countBefore || actualKind != request.ProcKind ||
                    !string.Equals(actual, request.Procedure, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBIDE did not recognize the inserted procedure.");
                string after = Code(module, countAfter);
                if (string.Equals(before, after, StringComparison.Ordinal))
                    throw new InvalidOperationException("VBIDE did not change the code module.");
                return new { Project = request.Project, Module = request.Module,
                    Procedure = actual, ProcKind = actualKind, BodyLine = bodyLine,
                    CountOfLines = countAfter, Sha256 = Hash(after), Code = after,
                    CompilationVerified = false };
            }
            catch
            {
                int insertedCount = (int)module.CountOfLines - countBefore;
                if (insertedCount > 0)
                    module.DeleteLines(countBefore + 1, insertedCount);
                throw;
            }
        }

        public object ReplaceProcedure(Request request)
        {
            string text = ValidateProcedureText(request);
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            dynamic module = GetModule(project, request.Module);
            int componentType = (int)module.Parent.Type;
            if (componentType != 1 && componentType != 2)
                throw new InvalidOperationException("replace_procedure targets a standard or class module.");
            int total = (int)module.CountOfLines;
            string before = Code(module, total);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int body = (int)module.ProcBodyLine[request.Procedure, request.ProcKind];
            int start = (int)module.ProcStartLine[request.Procedure, request.ProcKind];
            int count = (int)module.ProcCountLines[request.Procedure, request.ProcKind];
            if (body < start || start < 1 || count < 1 || start + count - 1 > total)
                throw new InvalidOperationException("VBIDE returned an invalid procedure range.");
            string oldDeclaration = (string)module.Lines[body, 1];
            string ending = request.ProcKind == 0 && Regex.IsMatch(oldDeclaration,
                @"\bFunction\s+", RegexOptions.IgnoreCase) ? "Function" :
                request.ProcKind == 0 ? "Sub" : "Property";
            int end = 0;
            for (int line = body; line < start + count; line++)
                if (Regex.IsMatch((string)module.Lines[line, 1],
                    @"^\s*End\s+" + ending + @"\s*$", RegexOptions.IgnoreCase)) end = line;
            if (end < body) throw new InvalidOperationException("The procedure's End statement could not be located safely.");
            string original = (string)module.Lines[body, end - body + 1];
            if (string.Equals(original.Replace("\r\n", "\n"), text, StringComparison.Ordinal))
                return new { Project = request.Project, Module = request.Module,
                    Procedure = request.Procedure, ProcKind = request.ProcKind,
                    BodyLine = body, CountOfLines = total, Sha256 = Hash(before),
                    Changed = false, CompilationVerified = false };
            int removed = end - body + 1;
            bool deleted = false;
            try
            {
                module.DeleteLines(body, removed);
                deleted = true;
                module.InsertLines(body, text.Replace("\n", "\r\n"));
                int newBody = (int)module.ProcBodyLine[request.Procedure, request.ProcKind];
                int actualKind = request.ProcKind;
                string actual = (string)module.ProcOfLine[newBody, ref actualKind];
                if (newBody != body || actualKind != request.ProcKind ||
                    !string.Equals(actual, request.Procedure, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBIDE did not recognize the replacement procedure.");
                string after = Code(module, (int)module.CountOfLines);
                return new { Project = request.Project, Module = request.Module,
                    Procedure = actual, ProcKind = actualKind, BodyLine = newBody,
                    CountOfLines = (int)module.CountOfLines, Sha256 = Hash(after),
                    Changed = true, CompilationVerified = false };
            }
            catch (Exception error)
            {
                if (deleted)
                {
                    try
                    {
                        int inserted = (int)module.CountOfLines - (total - removed);
                        if (inserted > 0) module.DeleteLines(body, inserted);
                        module.InsertLines(body, original);
                        if (!string.Equals(Code(module, (int)module.CountOfLines), before, StringComparison.Ordinal))
                            throw new InvalidOperationException("The original module text differs after rollback.");
                    }
                    catch (Exception rollbackError)
                    {
                        throw new InvalidOperationException("Procedure replacement failed and rollback needs inspection: " + rollbackError.Message, error);
                    }
                }
                throw;
            }
        }

        public object RemoveProcedure(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Procedure) ||
                !Regex.IsMatch(request.Procedure, @"^[A-Za-z][A-Za-z0-9_]{0,39}$") ||
                request.ProcKind < 0 || request.ProcKind > 3 ||
                string.IsNullOrWhiteSpace(request.ExpectedSha256))
                throw new ArgumentException("Procedure, ProcKind (0=Sub/Function, 1=Let, 2=Set, 3=Get) and ExpectedSha256 are required.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            dynamic module = GetModule(project, request.Module);
            int componentType = (int)module.Parent.Type;
            if (componentType != 1 && componentType != 2)
                throw new InvalidOperationException("remove_procedure targets a standard or class module.");
            int total = (int)module.CountOfLines;
            string before = Code(module, total);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int body = (int)module.ProcBodyLine[request.Procedure, request.ProcKind];
            int start = (int)module.ProcStartLine[request.Procedure, request.ProcKind];
            int count = (int)module.ProcCountLines[request.Procedure, request.ProcKind];
            if (body < start || start < 1 || count < 1 || start + count - 1 > total)
                throw new InvalidOperationException("VBIDE returned an invalid procedure range.");
            int actualKind = request.ProcKind;
            string actual = (string)module.ProcOfLine[body, ref actualKind];
            if (actualKind != request.ProcKind ||
                !string.Equals(actual, request.Procedure, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The procedure identity changed since it was read.");
            string declaration = (string)module.Lines[body, 1];
            string ending = request.ProcKind == 0 && Regex.IsMatch(declaration,
                @"\bFunction\s+", RegexOptions.IgnoreCase) ? "Function" :
                request.ProcKind == 0 ? "Sub" : "Property";
            int end = 0;
            for (int line = body; line < start + count; line++)
                if (Regex.IsMatch((string)module.Lines[line, 1],
                    @"^\s*End\s+" + ending + @"\s*$", RegexOptions.IgnoreCase)) end = line;
            if (end < body)
                throw new InvalidOperationException("The procedure's End statement could not be located safely.");
            int removed = end - body + 1;
            string original = (string)module.Lines[body, removed];
            bool deleted = false;
            try
            {
                module.DeleteLines(body, removed);
                deleted = true;
                int remainingBody = 0;
                try { remainingBody = (int)module.ProcBodyLine[request.Procedure, request.ProcKind]; }
                catch { }
                if (remainingBody > 0)
                    throw new InvalidOperationException("VBIDE still recognizes the removed procedure.");
                string after = Code(module, (int)module.CountOfLines);
                if (string.Equals(before, after, StringComparison.Ordinal))
                    throw new InvalidOperationException("VBIDE did not change the code module.");
                return new { Project = request.Project, Module = request.Module,
                    Procedure = actual, ProcKind = actualKind, RemovedStartLine = body,
                    RemovedLineCount = removed, CountOfLines = (int)module.CountOfLines,
                    Sha256 = Hash(after), Code = after, CompilationVerified = false };
            }
            catch (Exception error)
            {
                if (deleted)
                {
                    try
                    {
                        module.InsertLines(body, original);
                        if (!string.Equals(Code(module, (int)module.CountOfLines), before, StringComparison.Ordinal))
                            throw new InvalidOperationException("The original module text differs after rollback.");
                    }
                    catch (Exception rollbackError)
                    {
                        throw new InvalidOperationException("Procedure removal failed and rollback needs inspection: " + rollbackError.Message, error);
                    }
                }
                throw;
            }
        }

        public object InsertCodeFile(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                request.StartLine < 1 || string.IsNullOrWhiteSpace(request.Path) ||
                !Regex.IsMatch(request.Path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("ExpectedSha256, StartLine and an absolute Path are required.");
            string path = Path.GetFullPath(request.Path);
            var file = new FileInfo(path);
            if (!file.Exists) throw new FileNotFoundException("Code file not found.", path);
            if (file.Length == 0 || file.Length > 256 * 1024)
                throw new ArgumentException("The code file must contain 1 to 262144 bytes.");
            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length == 0 || bytes.Length > 256 * 1024)
                throw new ArgumentException("The code file changed size while it was read.");
            string sourceHash;
            using (var sha = SHA256.Create())
                sourceHash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            string encoding;
            string source = DecodeCodeFile(bytes, request.SourceEncoding, out encoding);
            if (source.IndexOf('\0') >= 0 || string.IsNullOrWhiteSpace(source))
                throw new ArgumentException("The selected file is not non-empty VBA source text.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            dynamic module = GetModule(project, request.Module);
            int beforeCount = (int)module.CountOfLines;
            string before = Code(module, beforeCount);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            if (request.StartLine > beforeCount + 1)
                throw new ArgumentOutOfRangeException("StartLine", "The insertion line is outside the module.");
            string normalized = source.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
            int insertedCount = 0;
            try
            {
                module.InsertLines(request.StartLine, normalized);
                insertedCount = (int)module.CountOfLines - beforeCount;
                string after = Code(module, (int)module.CountOfLines);
                if (insertedCount < 1 || string.Equals(after, before, StringComparison.Ordinal))
                    throw new InvalidOperationException("VBIDE did not insert the file content.");
                string insertedText = (string)module.Lines[request.StartLine, insertedCount];
                if (!string.Equals(NonAsciiCharacters(source), NonAsciiCharacters(insertedText), StringComparison.Ordinal))
                    throw new InvalidOperationException("VBE changed non-ASCII source characters while inserting the file.");
                return new { Project = request.Project, Module = request.Module,
                    Path = path, SourceByteCount = bytes.Length, SourceSha256 = sourceHash,
                    SourceEncoding = encoding, StartLine = request.StartLine,
                    InsertedLineCount = insertedCount, Sha256 = Hash(after),
                    CompilationVerified = false };
            }
            catch (Exception error)
            {
                if (insertedCount == 0)
                    insertedCount = Math.Max(0, (int)module.CountOfLines - beforeCount);
                if (insertedCount > 0)
                {
                    try
                    {
                        module.DeleteLines(request.StartLine, insertedCount);
                        if (!string.Equals(Code(module, (int)module.CountOfLines), before, StringComparison.Ordinal))
                            throw new InvalidOperationException("The original module text differs after rollback.");
                    }
                    catch (Exception rollbackError)
                    {
                        throw new InvalidOperationException("Code file insertion failed and rollback needs inspection: " + rollbackError.Message, error);
                    }
                }
                throw;
            }
        }

        public object InspectCodeFile(string suppliedPath)
        {
            if (string.IsNullOrWhiteSpace(suppliedPath) ||
                !Regex.IsMatch(suppliedPath, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("An absolute Path explicitly supplied by the user is required.");
            string path = Path.GetFullPath(suppliedPath);
            var file = new FileInfo(path);
            if (!file.Exists) throw new FileNotFoundException("Code file not found.", path);
            if (file.Length == 0 || file.Length > 256 * 1024)
                throw new ArgumentException("The code file must contain 1 to 262144 bytes.");
            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length == 0 || bytes.Length > 256 * 1024)
                throw new ArgumentException("The code file changed size while it was read.");
            string bom = bytes.Length >= 4 &&
                ((bytes[0] == 0xff && bytes[1] == 0xfe && bytes[2] == 0 && bytes[3] == 0) ||
                 (bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xfe && bytes[3] == 0xff)) ? "utf-32" :
                bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf ? "utf-8" :
                bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe ? "utf-16le" :
                bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff ? "utf-16be" : null;
            bool containsNonAscii = false;
            bool containsNul = false;
            foreach (byte value in bytes)
            {
                if (value >= 128) containsNonAscii = true;
                if (value == 0) containsNul = true;
            }
            bool? utf8Valid = null;
            if (bom == null || bom == "utf-8")
            {
                try
                {
                    int offset = bom == "utf-8" ? 3 : 0;
                    new UTF8Encoding(false, true).GetString(bytes, offset, bytes.Length - offset);
                    utf8Valid = true;
                }
                catch (DecoderFallbackException) { utf8Valid = false; }
            }
            string sha;
            using (var hash = SHA256.Create())
                sha = BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            return new { Path = path, ByteCount = bytes.Length, Sha256 = sha, Bom = bom,
                StrictUtf8Valid = utf8Valid, ContainsNonAscii = containsNonAscii,
                ContainsNulByte = containsNul,
                ExplicitEncodingRequired = bom == null && (containsNonAscii || containsNul),
                DefaultEncoding = bom ?? (!containsNonAscii && !containsNul ? "utf-8" : null),
                SystemAnsiCodePage = Encoding.Default.CodePage,
                SupportedSourceEncodings = new[] { "utf-8", "utf-16le", "utf-16be", "windows-1252", "system-ansi" },
                ContentIncluded = false };
        }

        private static string DecodeCodeFile(byte[] bytes, string requested, out string name)
        {
            string selected = string.IsNullOrWhiteSpace(requested) ? null : requested.Trim().ToLowerInvariant();
            if (selected != null && selected != "utf-8" && selected != "utf-16le" &&
                selected != "utf-16be" && selected != "windows-1252" && selected != "system-ansi")
                throw new ArgumentException("SourceEncoding must be utf-8, utf-16le, utf-16be, windows-1252 or system-ansi.");
            string bom = null;
            int offset = 0;
            if (bytes.Length >= 4 && ((bytes[0] == 0xff && bytes[1] == 0xfe && bytes[2] == 0 && bytes[3] == 0) ||
                (bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xfe && bytes[3] == 0xff)))
                throw new ArgumentException("UTF-32 source files are not supported; convert the file to UTF-8.");
            if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
            { bom = "utf-8"; offset = 3; }
            else if (bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe)
            { bom = "utf-16le"; offset = 2; }
            else if (bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff)
            { bom = "utf-16be"; offset = 2; }
            if (bom != null && selected != null && selected != bom)
                throw new ArgumentException("SourceEncoding conflicts with the file BOM.");
            if (bom == null && selected == null)
                for (int index = 0; index < bytes.Length; index++)
                    if (bytes[index] >= 128)
                        throw new ArgumentException("Non-ASCII source without a BOM is ambiguous. Inspect the file and specify SourceEncoding explicitly.");
            selected = selected ?? bom ?? "utf-8";
            Encoding decoder = selected == "utf-8" ? (Encoding)new UTF8Encoding(false, true) :
                selected == "utf-16le" ? new UnicodeEncoding(false, false, true) :
                selected == "utf-16be" ? new UnicodeEncoding(true, false, true) :
                Encoding.GetEncoding(selected == "windows-1252" ? 1252 : Encoding.Default.CodePage,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            string source;
            try { source = decoder.GetString(bytes, offset, bytes.Length - offset); }
            catch (DecoderFallbackException error)
            {
                throw new ArgumentException("The file is not valid " + selected +
                    ". Specify a different SourceEncoding only when that is the file's known encoding.", error);
            }
            byte[] roundTrip = decoder.GetBytes(source);
            if (roundTrip.Length != bytes.Length - offset)
                throw new ArgumentException("The selected encoding cannot reproduce the source bytes exactly.");
            for (int index = 0; index < roundTrip.Length; index++)
                if (roundTrip[index] != bytes[index + offset])
                    throw new ArgumentException("The selected encoding cannot reproduce the source bytes exactly.");
            name = selected;
            return source;
        }

        private static string NonAsciiCharacters(string source)
        {
            var result = new StringBuilder();
            foreach (char character in source)
                if (character > 127) result.Append(character);
            return result.ToString();
        }

        private static string ValidateProcedureText(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Procedure) ||
                !Regex.IsMatch(request.Procedure, @"^[A-Za-z][A-Za-z0-9_]{0,39}$") ||
                request.ProcKind < 0 || request.ProcKind > 3 ||
                string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                string.IsNullOrWhiteSpace(request.Text))
                throw new ArgumentException("Procedure, ProcKind (0=Sub/Function, 1=Let, 2=Set, 3=Get), ExpectedSha256 and Text are required.");
            string text = request.Text.Replace("\r\n", "\n").Replace('\r', '\n').Trim('\n');
            string[] lines = text.Split('\n');
            string kind = request.ProcKind == 0 ? @"(?:Sub|Function)" :
                "Property " + new[] { "", "Let", "Set", "Get" }[request.ProcKind];
            var declaration = Regex.Match(lines[0], @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*" +
                kind + @"\s+" + Regex.Escape(request.Procedure) + @"\s*\(", RegexOptions.IgnoreCase);
            string ending = request.ProcKind == 0 &&
                Regex.IsMatch(lines[0], @"\bFunction\s+", RegexOptions.IgnoreCase) ? "Function" :
                request.ProcKind == 0 ? "Sub" : "Property";
            if (!declaration.Success || !Regex.IsMatch(lines[lines.Length - 1],
                    @"^\s*End\s+" + ending + @"\s*$", RegexOptions.IgnoreCase))
                throw new ArgumentException("Text must contain one complete procedure with the requested declaration and End statement.");
            for (int line = 1; line < lines.Length - 1; line++)
                if (Regex.IsMatch(lines[line],
                    @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z]",
                    RegexOptions.IgnoreCase))
                    throw new ArgumentException("Text contains an additional procedure declaration.");
            return text;
        }

        private static int CountControls(IEnumerable nodes, string name)
        {
            int count = 0;
            foreach (dynamic node in nodes)
            {
                if (string.Equals((string)node.Kind, "Control", StringComparison.Ordinal) &&
                    string.Equals((string)node.Name, name, StringComparison.OrdinalIgnoreCase)) count++;
                count += CountControls((IEnumerable)node.Children, name);
            }
            return count;
        }

        public object Procedures(string projectName, string moduleName)
        {
            dynamic module = GetModule(GetProject(projectName), moduleName);
            int total = (int)module.CountOfLines;
            int declarations = (int)module.CountOfDeclarationLines;
            string code = Code(module, total);
            var result = new List<object>();
            int line = declarations + 1;
            while (line <= total)
            {
                int kind = 0;
                string name = (string)module.ProcOfLine[line, ref kind];
                if (string.IsNullOrWhiteSpace(name)) { line++; continue; }
                int start = (int)module.ProcStartLine[name, kind];
                int body = (int)module.ProcBodyLine[name, kind];
                int count = (int)module.ProcCountLines[name, kind];
                if (count < 1 || start < 1 || body < start || start + count - 1 > total)
                    throw new InvalidOperationException("VBIDE returned an invalid procedure range for " + name + ".");
                result.Add(new { Name = name, Kind = kind, StartLine = start,
                    BodyLine = body, CountLines = count, EndLine = start + count - 1,
                    Declaration = (string)module.Lines[body, 1] });
                line = Math.Max(line + 1, start + count);
            }
            return new { Project = projectName, Module = moduleName, Sha256 = Hash(code),
                CountOfLines = total, CountOfDeclarationLines = declarations,
                Procedures = result };
        }

        public object Find(Request request)
        {
            if (string.IsNullOrEmpty(request.Query) || request.Query.Length > 200)
                throw new ArgumentException("Query must contain 1 to 200 characters.");
            Regex wildcard = null;
            if (request.PatternSearch)
            {
                if (request.Query.Trim('*').Length == 0)
                    throw new ArgumentException("A wildcard pattern must contain more than asterisks.");
                string expression = Regex.Escape(request.Query).Replace(@"\*", ".*?").Replace(@"\?", ".");
                RegexOptions options = RegexOptions.CultureInvariant;
                if (!request.MatchCase) options |= RegexOptions.IgnoreCase;
                wildcard = new Regex(expression, options, TimeSpan.FromMilliseconds(200));
            }
            dynamic project = GetProject(request.Project);
            var results = new List<object>();
            var sourceVersions = new List<object>();
            var components = new List<dynamic>();
            foreach (dynamic component in project.VBComponents)
                if (string.IsNullOrWhiteSpace(request.Module) ||
                    string.Equals((string)component.Name, request.Module, StringComparison.OrdinalIgnoreCase))
                    components.Add(component);
            if (components.Count == 0) throw new InvalidOperationException("Module not found: " + request.Module);
            foreach (dynamic component in components)
            {
                dynamic module = component.CodeModule;
                int total = (int)module.CountOfLines;
                string code = Code(module, total);
                string name = (string)component.Name;
                string sha = Hash(code);
                sourceVersions.Add(new { Module = name, Sha256 = sha });
                if (total == 0) continue;
                string[] lines = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                if (lines.Length == total + 1 && lines[total].Length == 0)
                    Array.Resize(ref lines, total);
                if (lines.Length != total)
                {
                    lines = new string[total];
                    for (int line = 1; line <= total; line++)
                        lines[line - 1] = (string)module.Lines[line, 1];
                }
                for (int line = 0; line < lines.Length && results.Count <= 200; line++)
                {
                    string sourceLine = lines[line];
                    if (wildcard != null)
                    {
                        foreach (Match match in wildcard.Matches(sourceLine))
                        {
                            int afterMatch = match.Index + match.Length;
                            bool wholeMatch = !request.WholeWord ||
                                ((match.Index == 0 || !IdentifierChar(sourceLine[match.Index - 1])) &&
                                 (afterMatch == sourceLine.Length || !IdentifierChar(sourceLine[afterMatch])));
                            if (!wholeMatch) continue;
                            results.Add(new { Module = name, StartLine = line + 1,
                                StartColumn = match.Index + 1, EndLine = line + 1,
                                EndColumn = afterMatch, Text = sourceLine, Sha256 = sha });
                            if (results.Count > 200) break;
                        }
                        continue;
                    }
                    int offset = 0;
                    while (offset <= sourceLine.Length - request.Query.Length && results.Count <= 200)
                    {
                        int match = sourceLine.IndexOf(request.Query, offset,
                            request.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
                        if (match < 0) break;
                        int after = match + request.Query.Length;
                        bool word = !request.WholeWord ||
                            ((match == 0 || !IdentifierChar(sourceLine[match - 1])) &&
                             (after == sourceLine.Length || !IdentifierChar(sourceLine[after])));
                        if (word)
                            results.Add(new { Module = name, StartLine = line + 1,
                                StartColumn = match + 1, EndLine = line + 1, EndColumn = after,
                                Text = sourceLine, Sha256 = sha });
                        offset = match + 1;
                    }
                }
                if (results.Count > 200) break;
            }
            bool truncated = results.Count > 200;
            if (truncated) results.RemoveAt(200);
            return new { Project = request.Project, Query = request.Query,
                request.PatternSearch, request.WholeWord, request.MatchCase,
                Matches = results, SourceVersions = sourceVersions,
                Truncated = truncated };
        }

        private static bool IdentifierChar(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_';
        }

        public object SelectProcedure(Request request, VbeDebug debugger)
        {
            if (string.IsNullOrWhiteSpace(request.Procedure) ||
                request.ProcKind < 0 || request.ProcKind > 3)
                throw new ArgumentException("Procedure and ProcKind (0=Sub/Function, 1=Let, 2=Set, 3=Get) are required.");
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256))
                throw new ArgumentException("ExpectedSha256 from list_procedures or read_module is required.");
            dynamic module = GetModule(GetProject(request.Project), request.Module);
            string sha = Hash(Code(module, (int)module.CountOfLines));
            if (!string.Equals(sha, request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int body = (int)module.ProcBodyLine[request.Procedure, request.ProcKind];
            request.StartLine = body;
            return debugger.SelectCode(request);
        }

        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        private static dynamic GetModule(dynamic project, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Module is required.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                    return component.CodeModule;
            throw new InvalidOperationException("Module not found: " + name);
        }

        private static string Code(dynamic module, int count)
        {
            return count == 0 ? string.Empty : (string)module.Lines[1, count];
        }

        private static string Hash(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code)))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
