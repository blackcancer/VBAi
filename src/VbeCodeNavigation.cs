using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace CodexVBE
{
    internal sealed class VbeCodeNavigation
    {
        private readonly dynamic vbe;

        public VbeCodeNavigation(object vbe) { this.vbe = vbe; }

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
            if (request.MatchCase && request.PatternSearch)
                throw new ArgumentException("MatchCase and PatternSearch cannot both be true in VBIDE.Find.");
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
                int startLine = 1, startColumn = 1;
                while (startLine <= total && results.Count < 200)
                {
                    int endLine = total, endColumn = -1;
                    int searchedLine = startLine, searchedColumn = startColumn;
                    bool found = (bool)module.Find(request.Query, ref startLine, ref startColumn,
                        ref endLine, ref endColumn, request.WholeWord, request.MatchCase,
                        request.PatternSearch);
                    if (!found) break;
                    if (startLine < searchedLine ||
                        (startLine == searchedLine && startColumn < searchedColumn) ||
                        endLine < startLine || endLine > total || endColumn < 1)
                        throw new InvalidOperationException("VBIDE.Find returned an invalid or backwards range.");
                    results.Add(new { Module = name, StartLine = startLine,
                        StartColumn = startColumn, EndLine = endLine, EndColumn = endColumn,
                        Text = (string)module.Lines[startLine, 1], Sha256 = sha });
                    int nextLine = endLine, nextColumn = endColumn + 1;
                    if (nextLine == searchedLine && nextColumn <= searchedColumn)
                        nextColumn = searchedColumn + 1;
                    if (nextLine <= total)
                    {
                        string lastLine = (string)module.Lines[nextLine, 1];
                        if (nextColumn > lastLine.Length) { nextLine++; nextColumn = 1; }
                    }
                    startLine = nextLine;
                    startColumn = nextColumn;
                }
                if (results.Count >= 200) break;
            }
            return new { Project = request.Project, Query = request.Query,
                Matches = results, SourceVersions = sourceVersions,
                Truncated = results.Count >= 200 };
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
