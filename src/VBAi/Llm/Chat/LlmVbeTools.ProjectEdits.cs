using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Expose les opérations du chat qui inspectent ou modifient le code VBA.</summary>
    internal sealed partial class LlmVbeTools
    {

        /// <summary>Capture les sources du plan de renommage avant approbation et toute première écriture.</summary>
        /// <param name="request">Requête d’application contenant le projet et la version attendue.</param>
        /// <returns>Instantanés du code indexés par nom de module.</returns>
        private Dictionary<string, CodeSnapshot> ReadProcedureRenameBefore(Request request)
        {
            var previewRequest = new JavaScriptSerializer().Deserialize<Request>(json.Serialize(request));
            previewRequest.Command = request.Command == "apply_class_member_rename" ? "preview_class_member_rename" : "preview_procedure_rename";
            var response = Execute(previewRequest);
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            if (!(json.DeserializeObject(json.Serialize(response.Data)) is IDictionary<string, object> data) || !data.ContainsKey("ExpectedProjectVersion") ||
                !string.Equals(Convert.ToString(data["ExpectedProjectVersion"]), request.ExpectedProjectVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since the rename preview.");
            var before = new Dictionary<string, CodeSnapshot>(StringComparer.OrdinalIgnoreCase);
            foreach (object item in (object[])data["Edits"])
            {
                var edit = (IDictionary<string, object>)item;
                string module = (string)edit["Module"];
                before.Add(module, new CodeSnapshot { Code = (string)edit["Before"], Sha256 = (string)edit["ExpectedSha256"] });
            }
            return before;
        }

        /// <summary>Relit les modules même après une erreur partielle et publie seulement les différences observées.</summary>
        /// <param name="project">Projet dont les modules ont été inspectés.</param>
        /// <param name="before">Instantanés préalables indexés par module.</param>
        private void PublishProcedureRenameChanges(string project, Dictionary<string, CodeSnapshot> before)
        {
            foreach (var entry in before)
            {
                try
                {
                    var after = ReadCode(project, entry.Key);
                    if (!string.Equals(entry.Value.Sha256, after.Sha256, StringComparison.OrdinalIgnoreCase))
                        CodeEdited?.Invoke(new CodeChange(project, entry.Key, entry.Value.Code, entry.Value.Sha256,
                            after.Code, after.Sha256, CodeRollback.Lines(after.Code).Length));
                }
                catch (Exception) { WriteLog("Procedure rename diff readback failed for one inspected module."); }
            }
        }
    }
}
