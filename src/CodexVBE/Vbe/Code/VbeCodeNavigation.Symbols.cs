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
                    dynamic module = component.CodeModule;
                    int lineCount = (int)module.CountOfLines;
                    string source = lineCount == 0 ? "" : (string)module.Lines[1, lineCount];
                    foreach (var declaration in VbaDeclarationIndex.Read(source))
                        if (SymbolMatches(declaration.Name, request))
                            symbols.Add(new { Module = name, declaration.Name, declaration.Kind, declaration.Scope,
                                declaration.TypeName, declaration.Line, declaration.Column, declaration.Conditional,
                                ProcKind = DeclarationProcedureKind((System.Collections.IEnumerable)data.Procedures, declaration),
                                Sha256 = (string)data.Sha256, Declaration = CodeRollback.Lines(source)[declaration.Line - 1] });
                }
                catch (Exception ex) { errors.Add(new { Module = name, Error = ex.Message }); }
            }
            int size = request.Limit == 0 ? 50 : request.Limit;
            return new { Project = request.Project, Total = symbols.Count, Offset = request.Offset,
                Symbols = symbols.Skip(request.Offset).Take(size).ToArray(), HasMore = request.Offset + size < symbols.Count,
                Errors = errors, Scope = "Live VBIDE modules/procedures and syntactic VBA declarations (variables, constants, parameters, types and enum members). Conditional branches are marked; implicit variables, semantic binding and COM members are not inferred. Use list_reference_types/list_type_members for references." };
        }
        /// <summary>Associe une déclaration locale à l’accesseur ou à la procédure VBIDE qui la contient.</summary>
        private static int? DeclarationProcedureKind(System.Collections.IEnumerable procedures, VbaDeclarationIndex.Declaration declaration)
        {
            foreach (dynamic procedure in procedures)
                if (string.Equals((string)procedure.Name, declaration.Scope, StringComparison.OrdinalIgnoreCase) &&
                    declaration.Line >= (int)procedure.BodyLine && declaration.Line <= (int)procedure.EndLine)
                    return (int)procedure.Kind;
            return null;
        }
        private static bool SymbolMatches(string name, Request request)
        {
            var comparison = request.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return string.IsNullOrEmpty(request.Query) || (request.WholeWord ? string.Equals(name, request.Query, comparison) : name.IndexOf(request.Query, comparison) >= 0);
        }
    }
}
