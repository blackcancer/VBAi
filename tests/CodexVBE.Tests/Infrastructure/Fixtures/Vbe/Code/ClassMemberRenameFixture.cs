using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Sources et export réel temporaire de classe, sans hôte Office ni macros.</summary>
    public sealed class ClassMemberRenameFixture
    {
        public sealed class CodeSnapshot { public string Code { get; set; } }
        /// <summary>DTO public accessible au binder dynamique depuis la production.</summary>
        public sealed class ComponentSnapshot { public int Type { get; set; } public string Version { get; set; } }
        internal const string Target = "Option Explicit\r\nPrivate Function Calculate(ByVal x As Long) As Long\r\n    Calculate = x\r\nEnd Function\r\nPublic Sub UseIt()\r\n    Dim result As Long\r\n    result = Calculate(1) + Me.Calculate(2)\r\nEnd Sub";
        internal readonly Dictionary<string, string> Sources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        { { "CalcClass", Target }, { "Other", "Option Explicit\r\nPublic Sub Other(ByVal Calculate As Long)\r\nEnd Sub" } };
        internal readonly VbeCodeEdits Service;
        internal readonly List<string> Selectors = new List<string>(), ExportPaths = new List<string>();
        internal string MetadataVersion = "meta-v1", FailCommand, HiddenAttributes = "", HeaderOverride;
        internal int Mode = 2, Writes, Captures, Exports, TargetType = 2, NativeComponentType = 2;
        internal string NativeComponentVersion;
        internal bool WrongReadback, WrongExportBody, MissingExport, HiddenAfterWrite, OversizedExport;
        internal Action BeforeCatalogue;
        internal ClassMemberRenameFixture() { Service = new VbeCodeEdits(Execute); }
        internal ClassMemberRenameFixture(Func<string> exportPath) { Service = new VbeCodeEdits(Execute, exportPath); }
        internal static Request Request(string source = Target) => new Request { Project = "P", Module = "CalcClass", Query = "Calculate", NewName = "Compute",
            StartLine = 2, StartColumn = 18, ProcKind = 0, ExpectedMode = 2, ExpectedSha256 = VbaProcedureRename.Digest(source) };
        internal static string Export(string module, string source) => "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"" + module + "\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n" + source + "\r\n";
        private string ComponentVersion() => VbaProcedureRename.Digest(Sources["CalcClass"] + HiddenAttributes + HeaderOverride);
        private Response Execute(Request request)
        {
            Selectors.Add(request.Project);
            if (request.Command == FailCommand) return Response.Failure("injected " + request.Command);
            switch (request.Command)
            {
                case "project_properties": return Response.Success(new { Project = "P", Mode, Version = MetadataVersion });
                case "list_modules": Captures++; BeforeCatalogue?.Invoke(); return Response.Success(Sources.Keys.Select(name => new { Name = name, Type = name == "CalcClass" ? TargetType : 1 }).ToArray());
                case "read_module": return Response.Success(new CodeSnapshot { Code = Sources[request.Module] });
                case "component_properties": return Response.Success(new ComponentSnapshot { Type = NativeComponentType, Version = NativeComponentVersion ?? ComponentVersion() });
                case "export_component":
                    Exports++; ExportPaths.Add(request.Path);
                    if (request.ExpectedComponentVersion != ComponentVersion()) return Response.Failure("stale component version");
                    if (!MissingExport)
                    {
                        string text = Export(request.Module, WrongExportBody ? "Option Explicit" : Sources[request.Module]);
                        if (HeaderOverride != null) text = text.Replace("Attribute VB_Exposed = False", HeaderOverride);
                        text += HiddenAttributes;
                        if (HiddenAfterWrite && Writes > 0) text += "Attribute Compute.VB_UserMemId = 0\r\n";
                        if (OversizedExport) text = new string('x', 4 * 1024 * 1024 + 1);
                        File.WriteAllText(request.Path, text, Encoding.GetEncoding(Encoding.Default.CodePage));
                    }
                    return Response.Success(new { Path = request.Path });
                case "replace_lines":
                    Writes++; string before = Sources[request.Module];
                    if (request.ExpectedSha256 != VbaProcedureRename.Digest(before)) return Response.Failure("stale SHA");
                    Sources[request.Module] = WrongReadback ? "unexpected text" : request.Text;
                    Service.Record(request.Project, request.Module, before, Sources[request.Module]);
                    return Response.Success(new { Written = true });
                default: throw new InvalidOperationException("Unexpected command: " + request.Command);
            }
        }
        internal Request PreviewRequest()
        { var request = Request(Sources["CalcClass"]); dynamic preview = Service.PreviewClassMemberRename(request); request.ExpectedProjectVersion = (string)preview.ExpectedProjectVersion; return request; }
        internal static VbaProcedureRename.Plan Plan(string source, Request request = null, string other = "Option Explicit") => VbaClassMemberRename.Prepare("P",
            new[] { new VbaProcedureRename.ModuleSnapshot("CalcClass", 2, source), new VbaProcedureRename.ModuleSnapshot("Other", 1, other) }, request ?? Request(source));
    }
}