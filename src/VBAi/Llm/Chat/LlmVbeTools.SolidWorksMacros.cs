using System;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Applies project, privacy, editor, explicit-destination, and current edit-approval guards to native macro operations.</summary>
    internal sealed partial class LlmVbeTools
    {

        /// <summary>Recognizes the supported native macro creation and publication commands.</summary>
        /// <param name="name">Tool name to classify.</param>
        /// <returns><see langword="true"/> only for <c>create_solidworks_macro</c> and <c>publish_solidworks_macro</c>.</returns>
        private static bool IsSolidWorksMacroTool(string name) =>
            name == "create_solidworks_macro" || name == "publish_solidworks_macro";

        /// <summary>Freezes conversation grants and repeatedly revalidates mode, project, explicit path, legacy-editor, and VBE approval around one native macro operation.</summary>
        /// <param name="name">Create or publish tool name used by mode and project guards.</param>
        /// <param name="arguments">Original JSON arguments rechecked against the conversation's authorized project.</param>
        /// <param name="request">Parsed native request whose destination must remain explicitly user-authorized.</param>
        /// <param name="approved">Whether this concrete operation received AskEachTime approval.</param>
        /// <returns>Native outcome; if access changes after entry, the response marks uncertainty and disables retry.</returns>
        private async Task<Response> InvokeSolidWorksMacroAsync(string name, string arguments, Request request, bool approved)
        {
            string binding = BoundProject;
            string[] grants = readProjectGrants.ToArray();
            bool shared = sharedContextReadAllowed;
            request.RevalidateMacroAuthorization = live => {
                if (live) ValidateScope?.Invoke(); else ValidateCachedScope?.Invoke();
                GuardModeLocal(name);
                if (!SameProject(binding, BoundProject) || !readProjectGrants.SetEquals(grants) || sharedContextReadAllowed != shared)
                    throw new InvalidOperationException("Conversation project access changed during native macro creation/publication.");
                if (live) GuardProject(name, arguments);
                GuardLegacyEditorMutation(name);
                if (!IsExplicitUserPath(request.Path))
                    throw new InvalidOperationException("The destination path must remain explicitly authorized by the user.");
                if (settings.VbeEditApproval != "Automatic" && !(settings.VbeEditApproval == "AskEachTime" && approved))
                    throw new InvalidOperationException("VBE edit approval changed during native macro creation/publication.");
            };
            try
            {
                request.RevalidateMacroAuthorization(true);
                object result = await SolidWorksMacroNative(request);
                try { request.RevalidateMacroAuthorization(true); }
                catch
                {
                    var fields = Fields(result);
                    bool invoked = fields != null && ((fields.TryGetValue("MutationInvoked", out object m) && m is bool && (bool)m) ||
                        (fields.TryGetValue("CommandEntered", out object c) && c is bool && (bool)c));
                    return Response.Success(new { Available = false, Verified = false, MutationInvoked = invoked,
                        Uncertain = invoked, RetryAllowed = false,
                        Reason = "Project access changed during native macro creation/publication. Inspect locally; do not retry." });
                }
                return Response.Success(result);
            }
            finally { request.RevalidateMacroAuthorization = null; }
        }
    }
}
