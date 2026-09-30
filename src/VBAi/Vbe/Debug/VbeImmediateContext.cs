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
            var fields = serializer.DeserializeObject(serializer.Serialize(state)) as IDictionary<string, object>;
            object selected, path, resolved, module;
            if (string.IsNullOrWhiteSpace(project) || fields == null ||
                !fields.TryGetValue("Project", out resolved) || !(resolved is string) ||
                !fields.TryGetValue("SelectedProject", out selected) || !(selected is string) ||
                !fields.TryGetValue("ActiveModule", out module) || !(module is string) || string.IsNullOrWhiteSpace((string)module) ||
                !string.Equals((string)resolved, (string)selected, StringComparison.OrdinalIgnoreCase) ||
                (Path.IsPathRooted(project)
                    ? !(fields.TryGetValue("SelectedHostPath", out path) ||
                        (!fields.ContainsKey("SelectedHostPath") && VbeProjectHostPath.AllowsLegacyPath && fields.TryGetValue("SelectedProjectPath", out path))) || !(path is string) ||
                        !string.Equals(project, (string)path, StringComparison.OrdinalIgnoreCase)
                    : !string.Equals(project, (string)selected, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The requested project must be active in the VBE before Immediate execution. Select its code pane and read debug_state again.");
        }
    }
}
