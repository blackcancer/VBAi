using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Discovers explicitly annotated tests from plain source, without opening or executing a host.</summary>
    internal static class VbaTestDiscovery
    {

        /// <summary>Maintains the roles state for vba test discovery.</summary>
        private static readonly string[] Roles = { "TestMethod", "ModuleInitialize", "ModuleCleanup", "TestInitialize", "TestCleanup" };

        /// <summary>Maintains the annotation pattern state for vba test discovery.</summary>
        private static readonly Regex AnnotationPattern = new Regex(@"^\s*'\s*@(?<name>[A-Za-z]+)\b(?<argument>.*)$", RegexOptions.CultureInvariant);

        /// <summary>Maintains the quoted argument state for vba test discovery.</summary>
        private static readonly Regex QuotedArgument = new Regex("^\"(?:[^\"]|\"\")*\"$", RegexOptions.CultureInvariant);

        /// <summary>Owns the annotation state and operations.</summary>
        private sealed class Annotation
        {

            /// <summary>Maintains the name state for annotation.</summary>
            public string Name;

            /// <summary>Maintains the argument state for annotation.</summary>
            public string Argument;

            /// <summary>Maintains the line state for annotation.</summary>
            public int Line;
        }

        /// <summary>Handles discover for vba test discovery.</summary>
        /// <param name="project">vba test project snapshot that supplies the project for this operation.</param>
        /// <returns>vba test catalog produced by the operation for discover on vba test discovery.</returns>
        public static VbaTestCatalog Discover(VbaTestProjectSnapshot project)
        {
            if (project == null) throw new ArgumentNullException(nameof(project));
            var catalog = new VbaTestCatalog { Project = project };
            foreach (var source in (project.Modules ?? new VbaTestModuleSnapshot[0]).Where(x => x != null)
                .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            {
                var module = ReadModule(project, source, catalog.Diagnostics);
                if (module != null) catalog.Modules.Add(module);
            }
            foreach (var group in catalog.Modules.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1))
                foreach (var module in group)
                {
                    module.Diagnostic = Append(module.Diagnostic, "Duplicate module identity.");
                    catalog.Diagnostics.Add(module.Name + ": Duplicate module identity.");
                }
            return catalog;
        }

        /// <summary>Reads module for vba test discovery.</summary>
        /// <param name="project">vba test project snapshot that supplies the project for this operation.</param>
        /// <param name="snapshot">vba test module snapshot that supplies the snapshot for this operation.</param>
        /// <param name="diagnostics">list&lt;string&gt; that supplies the diagnostics for this operation.</param>
        /// <returns>vba test module produced by the operation for read module on vba test discovery.</returns>
        private static VbaTestModule ReadModule(VbaTestProjectSnapshot project, VbaTestModuleSnapshot snapshot, List<string> diagnostics)
        {
            var module = new VbaTestModule { Name = snapshot.Name ?? "" };
            var lines = (snapshot.Source ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var pending = new List<Annotation>();
            var declarations = new List<VbaTestDescriptor>();
            var procedureNames = new List<string>();
            int lastLine = 0, conditionalDepth = 0, moduleAnnotations = 0;
            bool inProcedure = false, sawProcedure = false, optionPrivate = false;
            foreach (var tokens in VbaDeclarationIndex.Statements(string.Join("\n", lines)))
            {
                if (tokens.Count == 0) continue;
                int firstLine = tokens[0].Line;
                if (!inProcedure)
                    ReadComments(lines, lastLine, firstLine - 1, pending, module, diagnostics, ref moduleAnnotations, sawProcedure, conditionalDepth);
                lastLine = tokens[tokens.Count - 1].Line;
                string Word(int index) => index < tokens.Count ? tokens[index].Text.ToLowerInvariant() : "";
                if (Word(0) == "#")
                {
                    ReportUnattached(pending, module, diagnostics);
                    pending.Clear();
                    if (Word(1) == "if") conditionalDepth++;
                    else if (Word(1) == "end" && Word(2) == "if") conditionalDepth = Math.Max(0, conditionalDepth - 1);
                    continue;
                }
                if (Word(0) == "end" && (Word(1) == "sub" || Word(1) == "function" || Word(1) == "property"))
                { inProcedure = false; pending.Clear(); continue; }
                if (inProcedure) continue;
                if (Word(0) == "option" && Word(1) == "private" && Word(2) == "module") optionPrivate = true;

                int member = 0;
                while (new[] { "public", "private", "friend", "global", "static" }.Contains(Word(member))) member++;
                bool external = Word(member) == "declare";
                if (external) { member++; if (Word(member) == "ptrsafe") member++; }
                bool procedure = Word(member) == "sub" || Word(member) == "function" || Word(member) == "property";
                if (procedure)
                {
                    sawProcedure = true;
                    if (!external) inProcedure = true;
                    int nameIndex = member + (Word(member) == "property" ? 2 : 1);
                    string name = nameIndex < tokens.Count ? tokens[nameIndex].Text : "<missing>";
                    procedureNames.Add(name);
                    var roles = pending.Where(x => Roles.Contains(x.Name, StringComparer.OrdinalIgnoreCase)).ToArray();
                    if (roles.Length > 0)
                    {
                        string kind = external ? "Declare" : CultureInfo.InvariantCulture.TextInfo.ToTitleCase(Word(member));
                        var test = new VbaTestDescriptor
                        {
                            Id = Identity(project.Id, module.Name, name), Module = module.Name,
                            Procedure = name, Line = firstLine, Kind = kind
                        };
                        if (roles.Length != 1) test.Diagnostic = "Duplicate or conflicting test/fixture annotations.";
                        bool fixture = !roles.Any(x => EqualsName(x.Name, "TestMethod"));
                        if (external || Word(member) == "property" || !ValidSignature(tokens, fixture))
                            test.Diagnostic = Append(test.Diagnostic, fixture
                                ? "Fixtures require an explicit Public parameterless Sub."
                                : "Tests require an explicit Public parameterless Sub or Function As Boolean.");
                        if (conditionalDepth > 0) test.Diagnostic = Append(test.Diagnostic, "Conditional compilation is unresolved.");
                        foreach (var role in roles.Where(x => x.Argument.Length != 0))
                            test.Diagnostic = Append(test.Diagnostic, "@" + role.Name + " does not accept an argument.");
                        ApplyMetadata(test, pending, fixture);
                        declarations.Add(test);
                        if (!fixture) module.Tests.Add(test);
                        else if (roles.Length == 1) SetFixture(module, roles[0].Name, test);
                        else module.Diagnostic = Append(module.Diagnostic, test.Diagnostic);
                    }
                    else ReportUnattached(pending, module, diagnostics);
                }
                else ReportUnattached(pending, module, diagnostics);
                pending.Clear();
            }
            if (!inProcedure)
                ReadComments(lines, lastLine, lines.Length, pending, module, diagnostics, ref moduleAnnotations, sawProcedure, conditionalDepth);
            ReportUnattached(pending, module, diagnostics);
            if (moduleAnnotations == 0 && declarations.Count == 0 && string.IsNullOrEmpty(module.Diagnostic)) return null;
            if (moduleAnnotations == 0) module.Diagnostic = Append(module.Diagnostic, "Missing @TestModule in the module declaration section.");
            if (moduleAnnotations > 1) module.Diagnostic = Append(module.Diagnostic, "Duplicate @TestModule annotation.");
            if (snapshot.ComponentType != 1) module.Diagnostic = Append(module.Diagnostic, "Tests are supported only in standard modules.");
            if (optionPrivate) module.Diagnostic = Append(module.Diagnostic, "Option Private Module is not supported for test invocation.");
            if (inProcedure && declarations.Count > 0) module.Diagnostic = Append(module.Diagnostic, "Unterminated procedure; discovery cannot establish a complete module.");
            foreach (var duplicate in procedureNames.GroupBy(x => x, StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1))
                foreach (var test in declarations.Where(x => EqualsName(x.Procedure, duplicate.Key)))
                    test.Diagnostic = Append(test.Diagnostic, "Duplicate procedure identity.");
            foreach (var fixture in new[] { module.ModuleInitialize, module.ModuleCleanup, module.TestInitialize, module.TestCleanup })
                if (fixture != null && !string.IsNullOrEmpty(fixture.Diagnostic))
                    module.Diagnostic = Append(module.Diagnostic, fixture.Procedure + ": " + fixture.Diagnostic);
            module.Tests = module.Tests.OrderBy(x => x.Procedure, StringComparer.OrdinalIgnoreCase).ThenBy(x => x.Line).ToList();
            if (!string.IsNullOrEmpty(module.Diagnostic)) diagnostics.Add(module.Name + ": " + module.Diagnostic);
            foreach (var test in module.Tests.Where(x => !string.IsNullOrEmpty(x.Diagnostic)))
                diagnostics.Add(module.Name + "." + test.Procedure + " (line " + test.Line + "): " + test.Diagnostic);
            return module;
        }

        /// <summary>Reads comments for vba test discovery.</summary>
        /// <param name="lines">string[] that supplies the lines for this operation.</param>
        /// <param name="start">int that supplies the start for this operation.</param>
        /// <param name="end">int that supplies the end for this operation.</param>
        /// <param name="pending">list&lt;annotation&gt; that supplies the pending for this operation.</param>
        /// <param name="module">vba test module that supplies the module for this operation.</param>
        /// <param name="diagnostics">list&lt;string&gt; that supplies the diagnostics for this operation.</param>
        /// <param name="moduleAnnotations">int that supplies the module annotations for this operation.</param>
        /// <param name="sawProcedure">Indicates whether saw procedure is enabled.</param>
        /// <param name="conditionalDepth">int that supplies the conditional depth for this operation.</param>
        private static void ReadComments(string[] lines, int start, int end, List<Annotation> pending, VbaTestModule module,
            List<string> diagnostics, ref int moduleAnnotations, bool sawProcedure, int conditionalDepth)
        {
            for (int i = start; i < end; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (!line.TrimStart().StartsWith("'", StringComparison.Ordinal))
                { ReportUnattached(pending, module, diagnostics); pending.Clear(); continue; }
                var match = AnnotationPattern.Match(line);
                if (!match.Success) continue;
                string name = match.Groups["name"].Value;
                if (EqualsName(name, "TestModule"))
                {
                    if (sawProcedure) { diagnostics.Add(module.Name + " (line " + (i + 1) + "): @TestModule must be in the declaration section."); continue; }
                    moduleAnnotations++;
                    if (match.Groups["argument"].Value.Trim().Length != 0)
                        module.Diagnostic = Append(module.Diagnostic, "@TestModule does not accept an argument.");
                    if (conditionalDepth > 0) module.Diagnostic = Append(module.Diagnostic, "Conditional @TestModule is unresolved.");
                }
                else if (Roles.Contains(name, StringComparer.OrdinalIgnoreCase) || EqualsName(name, "Ignore") || EqualsName(name, "TestCategory"))
                    pending.Add(new Annotation { Name = name, Argument = match.Groups["argument"].Value.Trim(), Line = i + 1 });
            }
        }

        /// <summary>Handles valid signature for vba test discovery.</summary>
        /// <param name="tokens">token&gt; that supplies the tokens for this operation.</param>
        /// <param name="fixture">Indicates whether fixture is enabled.</param>
        /// <returns>Boolean indicating the result of the check for valid signature on vba test discovery.</returns>
        private static bool ValidSignature(List<VbaDeclarationIndex.Token> tokens, bool fixture)
        {
            // Accept only the deliberately narrow version-one signature contract.
            string signature = string.Join(" ", tokens.Select(x => x.Text));
            string member = fixture ? "Sub" : "(?:Sub|Function)";
            var match = Regex.Match(signature, @"^Public (?<kind>" + member + @") \p{L}[\p{L}\p{N}_]* \( \)(?: As (?<type>Boolean))?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success) return false;
            return EqualsName(match.Groups["kind"].Value, "Sub") ? !match.Groups["type"].Success : match.Groups["type"].Success;
        }

        /// <summary>Handles apply metadata for vba test discovery.</summary>
        /// <param name="test">vba test descriptor that supplies the test for this operation.</param>
        /// <param name="annotations">list&lt;annotation&gt; that supplies the annotations for this operation.</param>
        /// <param name="fixture">Indicates whether fixture is enabled.</param>
        private static void ApplyMetadata(VbaTestDescriptor test, List<Annotation> annotations, bool fixture)
        {
            var categories = new List<string>();
            var ignores = annotations.Where(x => EqualsName(x.Name, "Ignore")).ToArray();
            if (ignores.Length > 1) test.Diagnostic = Append(test.Diagnostic, "Duplicate @Ignore annotation.");
            foreach (var annotation in annotations.Where(x => EqualsName(x.Name, "Ignore") || EqualsName(x.Name, "TestCategory")))
            {
                if (!QuotedArgument.IsMatch(annotation.Argument))
                { test.Diagnostic = Append(test.Diagnostic, "@" + annotation.Name + " requires a quoted nonempty string."); continue; }
                string value = annotation.Argument.Substring(1, annotation.Argument.Length - 2).Replace("\"\"", "\"");
                if (string.IsNullOrWhiteSpace(value))
                { test.Diagnostic = Append(test.Diagnostic, "@" + annotation.Name + " requires a quoted nonempty string."); continue; }
                if (fixture) test.Diagnostic = Append(test.Diagnostic, "Categories and @Ignore are not supported on fixtures.");
                if (EqualsName(annotation.Name, "Ignore")) test.IgnoreReason = value;
                else if (!categories.Contains(value, StringComparer.OrdinalIgnoreCase)) categories.Add(value);
            }
            test.Categories = categories.ToArray();
        }

        /// <summary>Sets fixture for vba test discovery.</summary>
        /// <param name="module">vba test module that supplies the module for this operation.</param>
        /// <param name="role">Text that supplies the role value. Use the format required by the calling operation.</param>
        /// <param name="fixture">vba test descriptor that supplies the fixture for this operation.</param>
        private static void SetFixture(VbaTestModule module, string role, VbaTestDescriptor fixture)
        {
            VbaTestDescriptor existing = null;
            if (EqualsName(role, "ModuleInitialize")) { existing = module.ModuleInitialize; if (existing == null) module.ModuleInitialize = fixture; }
            else if (EqualsName(role, "ModuleCleanup")) { existing = module.ModuleCleanup; if (existing == null) module.ModuleCleanup = fixture; }
            else if (EqualsName(role, "TestInitialize")) { existing = module.TestInitialize; if (existing == null) module.TestInitialize = fixture; }
            else if (EqualsName(role, "TestCleanup")) { existing = module.TestCleanup; if (existing == null) module.TestCleanup = fixture; }
            if (existing != null) module.Diagnostic = Append(module.Diagnostic, "Multiple @" + role + " fixtures.");
        }

        /// <summary>Handles report unattached for vba test discovery.</summary>
        /// <param name="annotations">list&lt;annotation&gt; that supplies the annotations for this operation.</param>
        /// <param name="module">vba test module that supplies the module for this operation.</param>
        /// <param name="diagnostics">list&lt;string&gt; that supplies the diagnostics for this operation.</param>
        private static void ReportUnattached(List<Annotation> annotations, VbaTestModule module, List<string> diagnostics)
        {
            foreach (var annotation in annotations)
                diagnostics.Add(module.Name + " (line " + annotation.Line + "): @" + annotation.Name + " is not attached to a procedure declaration.");
        }

        /// <summary>Handles identity for vba test discovery.</summary>
        /// <param name="project">Text that supplies the project value. Use the format required by the calling operation.</param>
        /// <param name="module">Text that supplies the module value. Use the format required by the calling operation.</param>
        /// <param name="procedure">Text that supplies the procedure value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for identity on vba test discovery.</returns>
        private static string Identity(string project, string module, string procedure)
        {
            // Source positions are navigation metadata; inserting lines must not rebind historical results.
            string input = string.Join("\0", (project ?? "").ToLowerInvariant(), module.ToLowerInvariant(), procedure.ToLowerInvariant());
            using (var hash = SHA256.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(input)).Select(x => x.ToString("x2", CultureInfo.InvariantCulture)));
        }

        /// <summary>Handles equals name for vba test discovery.</summary>
        /// <param name="left">Text that supplies the left value. Use the format required by the calling operation.</param>
        /// <param name="right">Text that supplies the right value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for equals name on vba test discovery.</returns>
        private static bool EqualsName(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

        /// <summary>Handles append for vba test discovery.</summary>
        /// <param name="current">Text that supplies the current value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for append on vba test discovery.</returns>
        private static string Append(string current, string message) => string.IsNullOrEmpty(current) ? message : current + " " + message;
    }
}
