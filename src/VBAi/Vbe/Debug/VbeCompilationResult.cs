using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Maps native command availability and diagnostic observation without claiming an unavailable compilation succeeded.</summary>
    internal static class VbeCompilationResult
    {
        internal static object Create(string project, object command, string diagnostic)
        {
            var json = new JavaScriptSerializer();
            var values = json.DeserializeObject(json.Serialize(command)) as IDictionary<string, object>;
            bool executed = values != null && values.TryGetValue("Executed", out var admitted) && admitted is bool flag && flag;
            bool available = executed;
            string capability = values != null && values.TryGetValue("Capability", out var state) ? Convert.ToString(state) : "Unknown";
            string reason = values != null && values.TryGetValue("Reason", out var why) ? Convert.ToString(why) : "The native Compile command was not executed; compilation was not verified.";
            return new
            {
                Project = project,
                Available = available,
                Compiled = executed && diagnostic == null,
                Diagnostic = executed ? diagnostic : reason,
                Verification = !executed ? "NativeCompile" + capability : diagnostic == null ? "NoNativeDiagnosticObserved" : "NativeDiagnosticCaptured",
                Command = command,
                NextRead = executed && diagnostic != null ? "Read debug_state to locate the selected token." : null
            };
        }
    }
}
