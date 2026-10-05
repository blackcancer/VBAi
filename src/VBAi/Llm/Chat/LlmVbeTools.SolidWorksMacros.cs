using System;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{
    internal sealed partial class LlmVbeTools
    {
        private static bool IsSolidWorksMacroTool(string name) =>
            name == "create_solidworks_macro" || name == "publish_solidworks_macro";

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
