using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Owns the vba coverage probe state and operations.</summary>
    internal sealed class VbaCoverageProbe
    {

        /// <summary>Gets or sets the index1 based.</summary>
        /// <value>Current index1 based exposed by vba coverage probe.</value>
        public int Index1Based { get; set; }

        /// <summary>Gets or sets the id.</summary>
        /// <value>Current id exposed by vba coverage probe.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the module.</summary>
        /// <value>Current module exposed by vba coverage probe.</value>
        public string Module { get; set; }

        /// <summary>Gets or sets the procedure.</summary>
        /// <value>Current procedure exposed by vba coverage probe.</value>
        public string Procedure { get; set; }

        /// <summary>Gets or sets the kind.</summary>
        /// <value>Current kind exposed by vba coverage probe.</value>
        public string Kind { get; set; }

        /// <summary>Gets or sets the original line.</summary>
        /// <value>Current original line exposed by vba coverage probe.</value>
        public int OriginalLine { get; set; }

        /// <summary>Gets or sets the original column.</summary>
        /// <value>Current original column exposed by vba coverage probe.</value>
        public int OriginalColumn { get; set; }

        /// <summary>Gets or sets the metric.</summary>
        /// <value>Current metric exposed by vba coverage probe.</value>
        public string Metric { get; set; } = "Procedure";
    }

    /// <summary>Owns the vba coverage module state and operations.</summary>
    internal sealed class VbaCoverageModule
    {

        /// <summary>Gets or sets the name.</summary>
        /// <value>Current name exposed by vba coverage module.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the component type.</summary>
        /// <value>Current component type exposed by vba coverage module.</value>
        public int ComponentType { get; set; }

        /// <summary>Gets or sets the original hash.</summary>
        /// <value>Current original hash exposed by vba coverage module.</value>
        public string OriginalHash { get; set; }

        /// <summary>Gets or sets the original source.</summary>
        /// <value>Current original source exposed by vba coverage module.</value>
        public string OriginalSource { get; set; }

        /// <summary>Gets or sets the instrumented source.</summary>
        /// <value>Current instrumented source exposed by vba coverage module.</value>
        public string InstrumentedSource { get; set; }

        /// <summary>Gets or sets the edits.</summary>
        /// <value>Current edits exposed by vba coverage module.</value>
        public List<VbaCoverageEdit> Edits { get; set; } = new List<VbaCoverageEdit>();
    }

    /// <summary>Owns the vba coverage edit state and operations.</summary>
    internal sealed class VbaCoverageEdit
    {

        /// <summary>Gets or sets the original line.</summary>
        /// <value>Current original line exposed by vba coverage edit.</value>
        public int OriginalLine { get; set; }

        /// <summary>Gets or sets the original column.</summary>
        /// <value>Current original column exposed by vba coverage edit.</value>
        public int OriginalColumn { get; set; }

        /// <summary>Gets or sets the text.</summary>
        /// <value>Current text exposed by vba coverage edit.</value>
        public string Text { get; set; }

        /// <summary>Gets or sets the is whole line.</summary>
        /// <value>Current is whole line exposed by vba coverage edit.</value>
        public bool IsWholeLine { get; set; }
    }

    /// <summary>Owns the vba coverage exclusion state and operations.</summary>
    internal sealed class VbaCoverageExclusion
    {

        /// <summary>Gets or sets the module.</summary>
        /// <value>Current module exposed by vba coverage exclusion.</value>
        public string Module { get; set; }

        /// <summary>Gets or sets the procedure.</summary>
        /// <value>Current procedure exposed by vba coverage exclusion.</value>
        public string Procedure { get; set; }

        /// <summary>Gets or sets the kind.</summary>
        /// <value>Current kind exposed by vba coverage exclusion.</value>
        public string Kind { get; set; }

        /// <summary>Gets or sets the original line.</summary>
        /// <value>Current original line exposed by vba coverage exclusion.</value>
        public int OriginalLine { get; set; }

        /// <summary>Gets or sets the reason.</summary>
        /// <value>Current reason exposed by vba coverage exclusion.</value>
        public string Reason { get; set; }

        /// <summary>Gets or sets the intentional.</summary>
        /// <value>Current intentional exposed by vba coverage exclusion.</value>
        public bool Intentional { get; set; }
    }

    /// <summary>Owns the vba coverage plan state and operations.</summary>
    internal sealed class VbaCoveragePlan
    {

        /// <summary>Gets or sets the original.</summary>
        /// <value>Current original exposed by vba coverage plan.</value>
        public string Original { get; set; }

        /// <summary>Gets or sets the revision.</summary>
        /// <value>Current revision exposed by vba coverage plan.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets the modules.</summary>
        /// <value>Current modules exposed by vba coverage plan.</value>
        public List<VbaCoverageModule> Modules { get; set; } = new List<VbaCoverageModule>();

        /// <summary>Gets or sets the probes.</summary>
        /// <value>Current probes exposed by vba coverage plan.</value>
        public List<VbaCoverageProbe> Probes { get; set; } = new List<VbaCoverageProbe>();

        /// <summary>Gets or sets the exclusions.</summary>
        /// <value>Current exclusions exposed by vba coverage plan.</value>
        public List<VbaCoverageExclusion> Exclusions { get; set; } = new List<VbaCoverageExclusion>();

        /// <summary>Gets or sets the diagnostics.</summary>
        /// <value>Current diagnostics exposed by vba coverage plan.</value>
        public List<string> Diagnostics { get; set; } = new List<string>();

        /// <summary>Gets or sets the denominator known.</summary>
        /// <value>Current denominator known exposed by vba coverage plan.</value>
        public bool DenominatorKnown { get; set; } = true;

        /// <summary>Gets or sets the eligible procedure count.</summary>
        /// <value>Current eligible procedure count exposed by vba coverage plan.</value>
        public int EligibleProcedureCount { get; set; }

        /// <summary>Gets the can instrument.</summary>
        /// <value>Current can instrument exposed by vba coverage plan.</value>
        public bool CanInstrument => DenominatorKnown && Diagnostics.Count == 0;

        /// <summary>Gets the statement coverage available.</summary>
        /// <value>Current statement coverage available exposed by vba coverage plan.</value>
        public bool StatementCoverageAvailable => false;

        /// <summary>Gets or sets the runtime source.</summary>
        /// <value>Current runtime source exposed by vba coverage plan.</value>
        public string RuntimeSource { get; set; }
    }

    /// <summary>Owns the vba coverage hit state and operations.</summary>
    internal sealed class VbaCoverageHit
    {

        /// <summary>Gets or sets the probe.</summary>
        /// <value>Current probe exposed by vba coverage hit.</value>
        public VbaCoverageProbe Probe { get; set; }

        /// <summary>Gets or sets the entered.</summary>
        /// <value>Current entered exposed by vba coverage hit.</value>
        public bool Entered { get; set; }
    }

    /// <summary>Owns the vba coverage report state and operations.</summary>
    internal sealed class VbaCoverageReport
    {

        /// <summary>Gets or sets the original.</summary>
        /// <value>Current original exposed by vba coverage report.</value>
        public string Original { get; set; }

        /// <summary>Gets or sets the revision.</summary>
        /// <value>Current revision exposed by vba coverage report.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets the metric.</summary>
        /// <value>Current metric exposed by vba coverage report.</value>
        public string Metric { get; set; } = "Procedure";

        /// <summary>Gets or sets the available.</summary>
        /// <value>Current available exposed by vba coverage report.</value>
        public bool Available { get; set; }

        /// <summary>Gets or sets the complete.</summary>
        /// <value>Current complete exposed by vba coverage report.</value>
        public bool Complete { get; set; }

        /// <summary>Gets or sets the denominator known.</summary>
        /// <value>Current denominator known exposed by vba coverage report.</value>
        public bool DenominatorKnown { get; set; }

        /// <summary>Gets or sets the eligible.</summary>
        /// <value>Current eligible exposed by vba coverage report.</value>
        public int? Eligible { get; set; }

        /// <summary>Gets or sets the hit.</summary>
        /// <value>Current hit exposed by vba coverage report.</value>
        public int? Hit { get; set; }

        /// <summary>Gets or sets the percent.</summary>
        /// <value>Current percent exposed by vba coverage report.</value>
        public double? Percent { get; set; }

        /// <summary>Gets the statement coverage available.</summary>
        /// <value>Current statement coverage available exposed by vba coverage report.</value>
        public bool StatementCoverageAvailable => false;

        /// <summary>Gets or sets the hits.</summary>
        /// <value>Current hits exposed by vba coverage report.</value>
        public List<VbaCoverageHit> Hits { get; set; } = new List<VbaCoverageHit>();

        /// <summary>Gets or sets the exclusions.</summary>
        /// <value>Current exclusions exposed by vba coverage report.</value>
        public List<VbaCoverageExclusion> Exclusions { get; set; } = new List<VbaCoverageExclusion>();

        /// <summary>Gets or sets the diagnostics.</summary>
        /// <value>Current diagnostics exposed by vba coverage report.</value>
        public List<string> Diagnostics { get; set; } = new List<string>();
    }

    /// <summary>Pure procedure-entry instrumentation, intended exclusively for a disposable project clone.</summary>
    internal static class VbaCoverageInstrumentation
    {

        /// <summary>Maintains the module name state for vba coverage instrumentation.</summary>
        internal const string ModuleName = "VBAiCoverageSupport";

        /// <summary>Maintains the reset procedure state for vba coverage instrumentation.</summary>
        internal const string ResetProcedure = "VBAiResetCoverageHits";

        /// <summary>Maintains the snapshot procedure state for vba coverage instrumentation.</summary>
        internal const string SnapshotProcedure = "VBAiReadCoverageHits";

        /// <summary>Maintains the hits variable state for vba coverage instrumentation.</summary>
        internal const string HitsVariable = "VBAiProcedureCoverageHits";

        /// <summary>Maintains the maximum probes state for vba coverage instrumentation.</summary>
        internal const int MaximumProbes = 16000;

        /// <summary>Maintains the name pattern state for vba coverage instrumentation.</summary>
        private static readonly Regex NamePattern = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}[$%&!#@^]?\z", RegexOptions.CultureInvariant);

        /// <summary>Maintains the test marker state for vba coverage instrumentation.</summary>
        private static readonly Regex TestMarker = new Regex(@"^\s*'\s*@TestModule(?:\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>Owns the insertion state and operations.</summary>
        private sealed class Insertion
        {

            /// <summary>Maintains the offset state for insertion.</summary>
            internal int Offset;

            /// <summary>Maintains the text state for insertion.</summary>
            internal string Text;
        }

        /// <summary>Creates  for vba coverage instrumentation.</summary>
        /// <param name="snapshot">vba test project snapshot that supplies the snapshot for this operation.</param>
        /// <returns>vba coverage plan produced by the operation for create on vba coverage instrumentation.</returns>
        public static VbaCoveragePlan Create(VbaTestProjectSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var plan = new VbaCoveragePlan { Original = snapshot.Selector ?? snapshot.Name, Revision = snapshot.Revision };
            var modules = snapshot.Modules ?? new VbaTestModuleSnapshot[0];
            foreach (var duplicate in modules.Where(module => module != null).GroupBy(module => module.Name ?? "", StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1))
                Diagnose(plan, duplicate.Key, 0, "Duplicate module identity makes coverage ambiguous.", true);
            foreach (var module in modules.OrderBy(module => module?.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (module == null) { Diagnose(plan, "<missing>", 0, "A module snapshot is missing.", true); continue; }
                if (string.Equals(module.Name, ModuleName, StringComparison.OrdinalIgnoreCase))
                { Diagnose(plan, module.Name, 0, "The coverage support module name is already occupied.", true); continue; }
                string source = module.Source ?? "";
                string normalized = source.Replace("\r\n", "\n").Replace('\r', '\n');
                var statements = VbaDeclarationIndex.Statements(normalized).Where(tokens => tokens.Count > 0).ToArray();
                var reserved = new HashSet<string>(new[] { ModuleName, ResetProcedure, SnapshotProcedure, HitsVariable }, StringComparer.OrdinalIgnoreCase);
                // Any existing use may bind differently after adding a Public runtime member,
                // including implicit variables in modules without Option Explicit.
                string[] sourceLines = normalized.Split('\n');
                var reservedUses = statements.SelectMany(tokens => tokens).SelectMany(token => IdentifierTokens(token, sourceLines)).Where(token =>
                    reserved.Contains(token.Text.TrimEnd('$', '%', '&', '!', '#', '@', '^')));
                foreach (var token in reservedUses.GroupBy(token => token.Text.TrimEnd('$', '%', '&', '!', '#', '@', '^'), StringComparer.OrdinalIgnoreCase).Select(group => group.First()))
                    Diagnose(plan, module.Name, token.Line, "Existing source uses a reserved coverage runtime identifier: " + token.Text + ".", false);
                if (reserved.Contains(module.Name ?? ""))
                    Diagnose(plan, module.Name, 0, "The module name conflicts with a reserved coverage runtime identifier.", false);
                int firstProcedureLine = statements.Where(IsHeader).Select(tokens => tokens[0].Line).DefaultIfEmpty(int.MaxValue).Min();
                bool testModule = normalized.Split('\n').Take(firstProcedureLine - 1).Any(line => TestMarker.IsMatch(line));
                if (testModule || string.Equals(module.Name, VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase))
                {
                    plan.Exclusions.Add(new VbaCoverageExclusion { Module = module.Name, Intentional = true,
                        Reason = testModule ? "Explicit @TestModule: tests and fixtures are excluded from production coverage." : "VBAi test framework support is excluded from production coverage." });
                    continue;
                }
                if (module.ComponentType != 1 && module.ComponentType != 2 && module.ComponentType != 3 && module.ComponentType != 100)
                { Diagnose(plan, module.Name, 0, "This component type cannot be instrumented safely.", true); continue; }
                if (!NamePattern.IsMatch(module.Name ?? ""))
                { Diagnose(plan, module.Name, 0, "The module identifier is unsupported.", true); continue; }

                var instrumented = new VbaCoverageModule { Name = module.Name, ComponentType = module.ComponentType,
                    OriginalSource = source, OriginalHash = Hash(source), InstrumentedSource = source };
                plan.Modules.Add(instrumented);
                var edits = new List<Insertion>();
                var offsets = LineOffsets(normalized);
                int conditionalDepth = 0;
                VbaCoverageProbe current = null;
                var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var members = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var tokens in statements)
                {
                    string Word(int index) => index < tokens.Count ? tokens[index].Text.ToLowerInvariant() : "";
                    if (Word(0) == "#")
                    {
                        if (Word(1) == "if") conditionalDepth++;
                        else if (Word(1) == "end" && Word(2) == "if")
                        {
                            if (conditionalDepth == 0) Diagnose(plan, module.Name, tokens[0].Line, "Unmatched conditional compilation directive.", true);
                            else conditionalDepth--;
                        }
                        else if ((Word(1) == "else" || Word(1) == "elseif") && conditionalDepth == 0)
                            Diagnose(plan, module.Name, tokens[0].Line, "Unmatched conditional compilation branch.", true);
                        else if (Word(1) != "const" && Word(1) != "else" && Word(1) != "elseif")
                            Diagnose(plan, module.Name, tokens[0].Line, "Unsupported conditional compilation directive.", true);
                        continue;
                    }
                    int at = HeaderMember(tokens);
                    if (at >= 0)
                    {
                        plan.EligibleProcedureCount++;
                        string kind = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(Word(at));
                        int nameIndex = at + 1;
                        if (kind == "Property")
                        {
                            if (Word(nameIndex) != "get" && Word(nameIndex) != "let" && Word(nameIndex) != "set")
                            { Diagnose(plan, module.Name, tokens[0].Line, "The Property accessor is ambiguous.", true); continue; }
                            kind += " " + CultureInfo.InvariantCulture.TextInfo.ToTitleCase(Word(nameIndex++));
                        }
                        if (nameIndex >= tokens.Count || !NamePattern.IsMatch(tokens[nameIndex].Text))
                        { Diagnose(plan, module.Name, tokens[0].Line, "The procedure identifier is ambiguous.", true); continue; }
                        string name = tokens[nameIndex].Text.TrimEnd('$', '%', '&', '!', '#', '@', '^');
                        if (string.Equals(name, ModuleName, StringComparison.OrdinalIgnoreCase))
                            Diagnose(plan, module.Name, tokens[0].Line, "A procedure shadows the coverage support module identifier.", false);
                        int depth = 0;
                        foreach (var token in tokens.Skip(nameIndex + 1))
                        {
                            if (token.Text == "(") depth++;
                            else if (token.Text == ")") depth--;
                            if (depth < 0) break;
                        }
                        if (depth != 0) Diagnose(plan, module.Name, tokens[0].Line, "The procedure signature has unmatched parentheses.", true);
                        var probe = new VbaCoverageProbe { Module = module.Name, Procedure = name, Kind = kind,
                            OriginalLine = tokens[nameIndex].Line, OriginalColumn = tokens[nameIndex].Column,
                            Id = Hash((snapshot.Id ?? "") + "\0" + module.Name.ToLowerInvariant() + "\0" + kind.ToLowerInvariant() + "\0" + name.ToLowerInvariant()) };
                        if (current != null) Diagnose(plan, module.Name, tokens[0].Line, "Nested or unterminated procedure declaration.", true);
                        current = probe;
                        bool duplicate = !identities.Add(kind + "." + name);
                        if (members.TryGetValue(name, out string previous) && (previous != "Property" || !kind.StartsWith("Property ", StringComparison.Ordinal))) duplicate = true;
                        members[name] = kind.StartsWith("Property ", StringComparison.Ordinal) ? "Property" : kind;
                        if (duplicate) Diagnose(plan, module.Name, tokens[0].Line, "Duplicate procedure identity: " + kind + " " + name + ".", true);
                        if (conditionalDepth != 0)
                        {
                            Diagnose(plan, module.Name, tokens[0].Line, "Conditional procedure declarations have an unresolved eligible denominator.", true);
                            plan.Exclusions.Add(new VbaCoverageExclusion { Module = module.Name, Procedure = name, Kind = kind,
                                OriginalLine = probe.OriginalLine, Reason = "Unresolved conditional declaration; this blocks project coverage.", Intentional = false });
                        }
                        probe.Index1Based = plan.Probes.Count + 1;
                        plan.Probes.Add(probe);
                        int offset = HeaderEndOffset(normalized, offsets, tokens);
                        if (offset < 0)
                        { Diagnose(plan, module.Name, tokens[0].Line, "The procedure header terminator cannot be mapped safely.", true); continue; }
                        string marker = ModuleName + "." + HitsVariable + "(" + probe.Index1Based.ToString(CultureInfo.InvariantCulture) + ") = True";
                        bool inline = normalized[offset - 1] == ':';
                        if (inline)
                        {
                            const string reason = "An inline procedure header would require replacing its declaration; preservation of hidden member attributes is not qualified.";
                            Diagnose(plan, module.Name, probe.OriginalLine, reason, false);
                            plan.Exclusions.Add(new VbaCoverageExclusion { Module = module.Name, Procedure = name, Kind = kind,
                                OriginalLine = probe.OriginalLine, Reason = reason + " This blocks project coverage.", Intentional = false });
                            continue;
                        }
                        string text = "    " + marker + "\n";
                        edits.Add(new Insertion { Offset = offset, Text = text });
                        int physicalLine = PhysicalLine(offsets, offset);
                        instrumented.Edits.Add(new VbaCoverageEdit { OriginalLine = physicalLine + 1,
                            OriginalColumn = offset - offsets[physicalLine] + 1, Text = text.TrimEnd('\r', '\n'), IsWholeLine = true });
                        continue;
                    }
                    if (Word(0) == "end" && (Word(1) == "sub" || Word(1) == "function" || Word(1) == "property"))
                    {
                        if (current == null || !current.Kind.StartsWith(CultureInfo.InvariantCulture.TextInfo.ToTitleCase(Word(1)), StringComparison.Ordinal))
                            Diagnose(plan, module.Name, tokens[0].Line, "Unmatched procedure terminator.", true);
                        if (conditionalDepth != 0) Diagnose(plan, module.Name, tokens[0].Line, "Conditional procedure terminators are ambiguous.", true);
                        current = null;
                    }
                    if (Word(0) == "attribute")
                        Diagnose(plan, module.Name, tokens[0].Line, "Exported Attribute statements are not supported as editable CodeModule source.", true);
                }
                if (current != null) Diagnose(plan, module.Name, current.OriginalLine, "Unterminated procedure.", true);
                if (conditionalDepth != 0) Diagnose(plan, module.Name, 0, "Unterminated conditional compilation block.", true);
                var rewritten = new StringBuilder(normalized);
                foreach (var edit in edits.OrderByDescending(edit => edit.Offset)) rewritten.Insert(edit.Offset, edit.Text);
                if (rewritten.ToString().Split('\n').Any(line => line.Length > 1023))
                    Diagnose(plan, module.Name, 0, "Instrumentation would exceed the VBA physical line-length limit.", false);
                instrumented.InstrumentedSource = source.Contains("\r\n") ? rewritten.ToString().Replace("\n", "\r\n") : rewritten.ToString();
            }
            if (plan.Probes.Count > MaximumProbes) Diagnose(plan, "<project>", 0, "The project exceeds the bounded coverage probe capacity.", false);
            plan.RuntimeSource = Runtime(plan.Probes.Count);
            return plan;
        }

        /// <summary>Reads  for vba coverage instrumentation.</summary>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <param name="native">object that supplies the native for this operation.</param>
        /// <param name="complete">Indicates whether complete is enabled.</param>
        /// <returns>vba coverage report produced by the operation for read on vba coverage instrumentation.</returns>
        public static VbaCoverageReport Read(VbaCoveragePlan plan, object native, bool complete = true)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (!plan.CanInstrument) throw new InvalidOperationException("The coverage plan has blocking diagnostics.");
            var hits = native as Array;
            int expected = Math.Max(1, plan.Probes.Count);
            if (hits == null || hits.Rank != 1 || hits.GetLowerBound(0) != 1 || hits.Length != expected)
                throw new InvalidOperationException("The coverage snapshot is not the exact one-based Boolean probe array.");
            for (int index = 1; index <= expected; index++)
                if (!(hits.GetValue(index) is bool)) throw new InvalidOperationException("The coverage snapshot contains a non-Boolean value.");
            if (plan.Probes.Count == 0 && (bool)hits.GetValue(1)) throw new InvalidOperationException("The empty coverage sentinel was changed.");
            var report = Report(plan);
            report.Available = true;
            report.Complete = complete;
            report.Hits = plan.Probes.Select(probe => new VbaCoverageHit { Probe = probe, Entered = (bool)hits.GetValue(probe.Index1Based) }).ToList();
            report.Hit = report.Hits.Count(hit => hit.Entered);
            report.Percent = complete && plan.EligibleProcedureCount > 0 ? 100d * report.Hit.Value / plan.EligibleProcedureCount : (double?)null;
            if (!complete) report.Diagnostics.Add("This run is incomplete; a complete project coverage percentage is unavailable.");
            return report;
        }

        /// <summary>Handles decode for vba coverage instrumentation.</summary>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <param name="native">object that supplies the native for this operation.</param>
        /// <param name="complete">Indicates whether complete is enabled.</param>
        /// <returns>vba coverage report produced by the operation for decode on vba coverage instrumentation.</returns>
        public static VbaCoverageReport Decode(VbaCoveragePlan plan, object native, bool complete = true) => Read(plan, native, complete);

        /// <summary>Handles unavailable for vba coverage instrumentation.</summary>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <param name="reason">Text that supplies the reason value. Use the format required by the calling operation.</param>
        /// <returns>vba coverage report produced by the operation for unavailable on vba coverage instrumentation.</returns>
        public static VbaCoverageReport Unavailable(VbaCoveragePlan plan, string reason)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            var report = Report(plan);
            if (!string.IsNullOrEmpty(reason)) report.Diagnostics.Add(reason);
            return report;
        }

        /// <summary>Handles report for vba coverage instrumentation.</summary>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <returns>vba coverage report produced by the operation for report on vba coverage instrumentation.</returns>
        private static VbaCoverageReport Report(VbaCoveragePlan plan) => new VbaCoverageReport { Original = plan.Original, Revision = plan.Revision,
            DenominatorKnown = plan.DenominatorKnown, Eligible = plan.DenominatorKnown ? plan.EligibleProcedureCount : (int?)null,
            Exclusions = plan.Exclusions.ToList(), Diagnostics = plan.Diagnostics.ToList() };

        /// <summary>Handles header member for vba coverage instrumentation.</summary>
        /// <param name="tokens">token&gt; that supplies the tokens for this operation.</param>
        /// <returns>int produced by the operation for header member on vba coverage instrumentation.</returns>
        private static int HeaderMember(List<VbaDeclarationIndex.Token> tokens)
        {
            int index = 0;
            while (index < tokens.Count && new[] { "Public", "Private", "Friend", "Global", "Static" }.Contains(tokens[index].Text, StringComparer.OrdinalIgnoreCase)) index++;
            return index < tokens.Count && new[] { "Sub", "Function", "Property" }.Contains(tokens[index].Text, StringComparer.OrdinalIgnoreCase) ? index : -1;
        }

        /// <summary>Determines whether header for vba coverage instrumentation.</summary>
        /// <param name="tokens">token&gt; that supplies the tokens for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is header on vba coverage instrumentation.</returns>
        private static bool IsHeader(List<VbaDeclarationIndex.Token> tokens) => HeaderMember(tokens) >= 0;

        /// <summary>Handles identifier tokens for vba coverage instrumentation.</summary>
        /// <param name="token">token that supplies the token for this operation.</param>
        /// <param name="lines">string[] that supplies the lines for this operation.</param>
        /// <returns>token&gt; produced by the operation for identifier tokens on vba coverage instrumentation.</returns>
        private static IEnumerable<VbaDeclarationIndex.Token> IdentifierTokens(VbaDeclarationIndex.Token token, string[] lines)
        {
            if (token.Text != "<literal>") { yield return token; yield break; }
            // Bracketed Excel expressions may resolve names/UDFs; they are not string literals.
            string line = lines[token.Line - 1];
            int start = token.Column - 1;
            if (start >= line.Length || line[start] != '[') yield break;
            int end = line.IndexOf(']', start + 1);
            string expression = line.Substring(start + 1, (end < 0 ? line.Length : end) - start - 1);
            foreach (Match match in Regex.Matches(expression, @"\p{L}[\p{L}\p{N}_]*[$%&!#@^]?", RegexOptions.CultureInvariant))
                yield return new VbaDeclarationIndex.Token { Text = match.Value, Line = token.Line, Column = token.Column + 1 + match.Index };
        }

        /// <summary>Handles line offsets for vba coverage instrumentation.</summary>
        /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
        /// <returns>int[] produced by the operation for line offsets on vba coverage instrumentation.</returns>
        private static int[] LineOffsets(string source)
        {
            var offsets = new List<int> { 0 };
            for (int i = 0; i < source.Length; i++) if (source[i] == '\n') offsets.Add(i + 1);
            return offsets.ToArray();
        }

        /// <summary>Handles header end offset for vba coverage instrumentation.</summary>
        /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
        /// <param name="offsets">int[] that supplies the offsets for this operation.</param>
        /// <param name="tokens">token&gt; that supplies the tokens for this operation.</param>
        /// <returns>int produced by the operation for header end offset on vba coverage instrumentation.</returns>
        private static int HeaderEndOffset(string source, int[] offsets, List<VbaDeclarationIndex.Token> tokens)
        {
            var last = tokens[tokens.Count - 1];
            if (last.Text == "<literal>") return -1;
            int offset = offsets[last.Line - 1] + last.Column - 1 + last.Text.Length;
            for (; offset < source.Length; offset++)
            {
                if (source[offset] == ':') return offset + 1;
                if (source[offset] == '\n') return offset + 1;
                if (source[offset] == '\'')
                {
                    int end = source.IndexOf('\n', offset);
                    return end < 0 ? -1 : end + 1;
                }
                if (!char.IsWhiteSpace(source[offset])) return -1;
            }
            return -1;
        }

        /// <summary>Handles physical line for vba coverage instrumentation.</summary>
        /// <param name="offsets">int[] that supplies the offsets for this operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        /// <returns>int produced by the operation for physical line on vba coverage instrumentation.</returns>
        internal static int PhysicalLine(int[] offsets, int offset)
        {
            int physicalLine = Array.BinarySearch(offsets, offset);
            return physicalLine < 0 ? ~physicalLine - 1 : physicalLine;
        }

        /// <summary>Runs time for vba coverage instrumentation.</summary>
        /// <param name="count">int that supplies the count for this operation.</param>
        /// <returns>Text produced by the operation for runtime on vba coverage instrumentation.</returns>
        private static string Runtime(int count)
        {
            return "Option Explicit\r\n' VBAi procedure coverage support version 1; disposable clone only.\r\n"
                + "Public " + HitsVariable + "(1 To " + Math.Max(1, count).ToString(CultureInfo.InvariantCulture) + ") As Boolean\r\n"
                + "Public Function " + ResetProcedure + "(Optional ByVal ignoredHostArgument1 As Variant, Optional ByVal ignoredHostArgument2 As Variant) As Boolean\r\n    Dim index As Long\r\n    For index = 1 To " + count.ToString(CultureInfo.InvariantCulture)
                + "\r\n        " + HitsVariable + "(index) = False\r\n    Next index\r\n    " + ResetProcedure + " = True\r\nEnd Function\r\n"
                + "Public Function " + SnapshotProcedure + "(Optional ByVal ignoredHostArgument1 As Variant, Optional ByVal ignoredHostArgument2 As Variant) As Variant\r\n    " + SnapshotProcedure + " = " + HitsVariable + "\r\nEnd Function\r\n";
        }

        /// <summary>Handles diagnose for vba coverage instrumentation.</summary>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <param name="module">Text that supplies the module value. Use the format required by the calling operation.</param>
        /// <param name="line">int that supplies the line for this operation.</param>
        /// <param name="reason">Text that supplies the reason value. Use the format required by the calling operation.</param>
        /// <param name="unknownDenominator">Indicates whether unknown denominator is enabled.</param>
        private static void Diagnose(VbaCoveragePlan plan, string module, int line, string reason, bool unknownDenominator)
        {
            plan.Diagnostics.Add((module ?? "<missing>") + (line > 0 ? " (line " + line.ToString(CultureInfo.InvariantCulture) + ")" : "") + ": " + reason);
            if (unknownDenominator) plan.DenominatorKnown = false;
        }

        /// <summary>Determines whether it has h for vba coverage instrumentation.</summary>
        /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for hash on vba coverage instrumentation.</returns>
        private static string Hash(string source)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(source ?? ""))).Replace("-", ""); }
    }
}
