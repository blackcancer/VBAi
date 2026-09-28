using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Sonde de projet complète, gardes SHA natives et journalisation existante simulées.</summary>
    public sealed class ProcedureRenameWorkflowFixture
    {
        /// <summary>Réponse de lecture accessible au binder dynamique du service existant.</summary>
        public sealed class CodeSnapshot { public string Code { get; set; } }
        internal readonly Dictionary<string, string> Sources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        { { "MathModule", ProcedureRenameMatrix.Target }, { "Caller", ProcedureRenameMatrix.Caller } };
        internal readonly VbeCodeEdits Service;
        internal readonly List<string> Selectors = new List<string>();
        internal string MetadataVersion = "metadata-v1", FailCommand, FailedModule;
        internal int Mode = 2, Writes, Captures;
        internal bool WrongReadback;
        internal Action BeforeCatalogue;

        /// <summary>Configure le service avec un exécutant qui enregistre les mutations relues par le workflow.</summary>
        internal ProcedureRenameWorkflowFixture()
        { Service = new VbeCodeEdits(Execute); }

        /// <summary>Exécute seulement les commandes nécessaires au plan et à l'historique par module.</summary>
        private Response Execute(Request request)
        {
            Selectors.Add(request.Project);
            if (request.Command == FailCommand) return Response.Failure("injected " + request.Command);
            switch (request.Command)
            {
                case "project_properties": return Response.Success(new { Project = "P", Mode, Version = MetadataVersion });
                case "list_modules":
                    Captures++; BeforeCatalogue?.Invoke();
                    return Response.Success(Sources.Keys.Select(name => new { Name = name, Type = name == "MathModule" ? 1 : 2 }).ToArray());
                case "read_module": return Response.Success(new CodeSnapshot { Code = Sources[request.Module] });
                case "replace_lines":
                    Writes++;
                    if (request.Module == FailedModule) return Response.Failure("injected write failure");
                    string before = Sources[request.Module];
                    if (!string.Equals(VbaProcedureRename.Digest(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase)) return Response.Failure("stale SHA");
                    Sources[request.Module] = WrongReadback ? "unexpected text" : request.Text;
                    Service.Record(request.Project, request.Module, before, Sources[request.Module]);
                    return Response.Success(new { Written = true });
                default: throw new InvalidOperationException("Unexpected command: " + request.Command);
            }
        }

        /// <summary>Prépare une requête d'application à partir de la version réellement prévisualisée.</summary>
        internal Request PreviewRequest()
        {
            var request = ProcedureRenameMatrix.Request(); dynamic preview = Service.PreviewProcedureRename(request);
            request.ExpectedProjectVersion = (string)preview.ExpectedProjectVersion; return request;
        }
    }
}
