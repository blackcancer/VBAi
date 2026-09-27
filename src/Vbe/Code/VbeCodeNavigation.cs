using System;
using System.Collections.Generic;
using System.Collections;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed class VbeCodeNavigation
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
            if (request.PatternSearch)
                throw new InvalidOperationException("Wildcard search is not yet supported by find_code.");
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
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Project is required.");
            dynamic found = null;
            foreach (dynamic project in vbe.VBProjects)
                if (string.Equals((string)project.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (found != null) throw new InvalidOperationException("Project name is ambiguous.");
                    found = project;
                }
            if (found == null) throw new InvalidOperationException("Project not found: " + name);
            return found;
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
