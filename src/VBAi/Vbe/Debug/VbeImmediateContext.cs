using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Checks the native project selected for a global Immediate execution.</summary>
    internal static class VbeImmediateContext
    {

        /// <summary>Reads fresh state on the owning STA immediately before submitting the native Enter message.</summary>
        /// <param name="project">Approved exact project selector.</param>
        /// <param name="expectedMode">Approved project mode.</param>
        /// <param name="execute">Owner-thread session boundary.</param>
        internal static void RequireCurrent(string project, int expectedMode, Func<Request, Response> execute)
        {
            var state = execute(new Request { Command = "debug_state", Project = project });
            if (!state.Ok) throw new InvalidOperationException(state.Error);
            if ((int)((dynamic)state.Data).Mode != expectedMode)
                throw new InvalidOperationException("Project mode changed before Immediate execution; Enter was not sent.");
            RequireProject(project, state.Data);
        }

        /// <summary>Refuses execution when the active native code pane does not belong to the requested project.</summary>
        /// <param name="project">Exact project selector resolved by debug_state.</param>
        /// <param name="state">Unfiltered native debug_state response, whose selection is checked by COM identity.</param>
        internal static void RequireProject(string project, object state)
        {
            var serializer = new JavaScriptSerializer();
            if (string.IsNullOrWhiteSpace(project) || !(serializer.DeserializeObject(serializer.Serialize(state)) is IDictionary<string, object> fields) ||
                !fields.TryGetValue("Project", out object resolved) || !(resolved is string v) ||
                !fields.TryGetValue("SelectedProject", out object selected) || !(selected is string v1) ||
                !fields.TryGetValue("ActiveModule", out object module) || !(module is string v2) || string.IsNullOrWhiteSpace(v2) ||
                !string.Equals(v, v1, StringComparison.OrdinalIgnoreCase) ||
                (Path.IsPathRooted(project)
                    ? !(fields.TryGetValue("SelectedHostPath", out object path) ||
                        (!fields.ContainsKey("SelectedHostPath") && VbeProjectHostPath.AllowsLegacyPath && fields.TryGetValue("SelectedProjectPath", out path))) || !(path is string v3) ||
                        !string.Equals(project, v3, StringComparison.OrdinalIgnoreCase)
                    : !string.Equals(project, v1, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The requested project must be active in the VBE before Immediate execution. Select its code pane and read debug_state again.");
        }
    }
}
