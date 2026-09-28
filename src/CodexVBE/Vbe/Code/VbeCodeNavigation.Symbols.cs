using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeCodeNavigation
    {
        public object ProjectSymbols(Request request)
        {
            if (request.Offset < 0 || request.Limit < 0 || request.Limit > 200 || (request.Query ?? "").Length > 200)
                throw new ArgumentException("Invalid symbol search bounds.");
            dynamic project = GetProject(request.Project);
            var symbols = new List<object>(); var errors = new List<object>();
            foreach (dynamic component in project.VBComponents)
            {
                string name = (string)component.Name;
                if (!string.IsNullOrEmpty(request.Module) && !string.Equals(name, request.Module, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    dynamic data = Procedures(request.Project, name);
                    if (SymbolMatches(name, request)) symbols.Add(new { Module = name, Name = name, Kind = "Module", Line = 1, Sha256 = (string)data.Sha256, Declaration = name });
                    foreach (dynamic procedure in (IEnumerable)data.Procedures)
                        if (SymbolMatches((string)procedure.Name, request))
                            symbols.Add(new { Module = name, Name = (string)procedure.Name, Kind = "Procedure", ProcKind = (int)procedure.Kind,
                                Line = (int)procedure.BodyLine, Sha256 = (string)data.Sha256, Declaration = (string)procedure.Declaration });
                }
                catch (Exception ex) { errors.Add(new { Module = name, Error = ex.Message }); }
            }
            int size = request.Limit == 0 ? 50 : request.Limit;
            return new { Project = request.Project, Total = symbols.Count, Offset = request.Offset,
                Symbols = symbols.Skip(request.Offset).Take(size).ToArray(), HasMore = request.Offset + size < symbols.Count,
                Errors = errors, Scope = "Live VBIDE modules and Sub/Function/Property definitions; local variables and COM reference members are not included. Use list_reference_types/list_type_members for references." };
        }
        private static bool SymbolMatches(string name, Request request)
        {
            var comparison = request.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return string.IsNullOrEmpty(request.Query) || (request.WholeWord ? string.Equals(name, request.Query, comparison) : name.IndexOf(request.Query, comparison) >= 0);
        }
    }
}
