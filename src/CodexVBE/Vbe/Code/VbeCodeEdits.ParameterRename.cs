using System;

namespace CodexVBE
{
    /// <summary>Partie de l’éditeur transactionnel consacrée au renommage de paramètres et d’appels nommés.</summary>
    internal sealed partial class VbeCodeEdits
    {
        /// <summary>Prévisualise ou applique le renommage d'un paramètre privé avec mise à jour des appels nommés locaux.</summary>
        /// <param name="request">Projet, module, procédure, déclaration et SHA inspectés.</param>
        /// <param name="preview">Vrai pour calculer le diff sans écriture.</param>
        /// <returns>Diff préparé ou résultat relu de l'édition du module.</returns>
        internal object RenameParameter(Request request, bool preview)
        {
            string before = Read(request.Project, request.Module);
            Check(before, request.ExpectedSha256);
            Response component = execute(new Request { Command = "component_properties", Project = request.Project, Module = request.Module });
            if (!component.Ok) throw new InvalidOperationException(component.Error);
            if ((int)((dynamic)component.Data).Type != 1)
                throw new InvalidOperationException("Parameter renaming currently requires a standard module; class, form and host callbacks need a wider binding plan.");
            Response catalog = execute(new Request { Command = "list_procedures", Project = request.Project, Module = request.Module });
            if (!catalog.Ok) throw new InvalidOperationException(catalog.Error);
            dynamic data = catalog.Data;
            if (!string.Equals((string)data.Sha256, request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed before resolving the parameter's procedure.");
            dynamic selected = null;
            foreach (dynamic procedure in data.Procedures)
                if ((int)procedure.Kind == request.ProcKind && string.Equals((string)procedure.Name, request.Procedure, StringComparison.OrdinalIgnoreCase))
                { if (selected != null) throw new InvalidOperationException("The procedure is ambiguous."); selected = procedure; }
            if (selected == null) throw new InvalidOperationException("The exact procedure is absent.");
            string after = VbaParameterRename.Transform(before, request, (int)selected.BodyLine, (int)selected.EndLine);
            if (preview) return new { request.Project, request.Module, request.Procedure, Before = before, After = after,
                ExpectedSha256 = Hash(before), Changed = before != after,
                Scope = "Private standard-module parameter, local uses and named arguments of direct or module-qualified calls in this module. Public interfaces, classes, callbacks and conditional binding are refused." };
            if (request.ExpectedMode != 2) throw new ArgumentException("ExpectedMode=2 is required to rename a parameter.");
            Response state = execute(new Request { Command = "debug_state", Project = request.Project });
            if (!state.Ok || (int)((dynamic)state.Data).Mode != 2) throw new InvalidOperationException("Renaming requires design mode.");
            return Write(request.Project, request.Module, before, after);
        }
    }
}
