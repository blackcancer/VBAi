using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Capture toutes les sources du plan de renommage avant l'approbation et la première écriture.</summary>
        private Dictionary<string, CodeSnapshot> ReadProcedureRenameBefore(Request request)
        {
            var previewRequest = new JavaScriptSerializer().Deserialize<Request>(json.Serialize(request));
            previewRequest.Command = request.Command == "apply_class_member_rename" ? "preview_class_member_rename" : "preview_procedure_rename";
            var response = Execute(previewRequest);
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
            if (data == null || !data.ContainsKey("ExpectedProjectVersion") ||
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
