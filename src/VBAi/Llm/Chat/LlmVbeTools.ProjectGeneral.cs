using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBAi
{
    /// <summary>Routes native General operations through the ordinary guarded asynchronous tool pipeline.</summary>
    internal sealed partial class LlmVbeTools
    {
        internal Func<Request, bool, Task<object>> ProjectGeneralNative;

        private static bool IsProjectGeneralTool(string name) => name == "read_project_general" || name == "set_project_general";

        private static void ValidateProjectGeneralRequest(Request request)
        {
            if (request.ExpectedMode != 2)
                throw new ArgumentException("ExpectedMode=2 is required for native General properties.");
            if (request.Command != "set_project_general") return;
            if (request.Property == "HelpFile")
            {
                if (!(request.Value is string file) || string.IsNullOrWhiteSpace(file) ||
                    !Regex.IsMatch(file, @"\A(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])") ||
                    !Path.GetExtension(file).Equals(".chm", StringComparison.OrdinalIgnoreCase) || file.IndexOf('\0') >= 0)
                    throw new ArgumentException("HelpFile must be an absolute CHM path.");
                return;
            }
            if (request.Property != "HelpContextID")
                throw new ArgumentException("Only HelpContextID and HelpFile are supported by set_project_general.");
            if (!(request.Value is int) && !(request.Value is long))
                throw new ArgumentException("HelpContextID must be a nonnegative Int32 integer.");
            long value = Convert.ToInt64(request.Value);
            if (value < 0 || value > int.MaxValue)
                throw new ArgumentException("HelpContextID must be a nonnegative Int32 integer.");
        }

        private async Task<Response> InvokeProjectGeneralAsync(string name, string arguments, Request request, bool editApproved)
        {
            bool write = name == "set_project_general";
            string originalBinding = BoundProject;
            string[] originalReadGrants = readProjectGrants.ToArray();
            request.RevalidateProjectPropertyAuthorization = live => {
                if (live) ValidateScope?.Invoke(); else ValidateCachedScope?.Invoke();
                GuardModeLocal(name);
                if (!string.Equals(originalBinding, BoundProject, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The conversation project binding changed during native General dispatch.");
                if (live || write) GuardProject(name, arguments);
                else if (!readProjectGrants.SetEquals(originalReadGrants))
                    throw new InvalidOperationException("Project read grants changed during native General inspection.");
                GuardLegacyEditorMutation(name);
                if (write && settings.VbeEditApproval != "Automatic" &&
                    !(settings.VbeEditApproval == "AskEachTime" && editApproved))
                    throw new InvalidOperationException("VBE edit policy changed during native General dispatch.");
            };
            try
            {
                request.RevalidateProjectPropertyAuthorization(true);
                object result = await ProjectGeneralNative(request, write);
                // Read authorization can change while the modal native command is pending.
                // Revalidate before any General metadata is transmitted to the provider.
                try { request.RevalidateProjectPropertyAuthorization(true); }
                catch
                {
                    if (!write) return Response.Failure("Project access changed during native General inspection.");
                    var fields = Fields(result);
                    bool invoked = fields != null && fields.TryGetValue("MutationInvoked", out object mutation) && mutation is bool && (bool)mutation;
                    bool uncertain = invoked || (fields != null && fields.TryGetValue("Uncertain", out object pending) && pending is bool && (bool)pending);
                    return Response.Success(new { Available = false, MutationInvoked = invoked,
                        Uncertain = uncertain, RetryAllowed = false,
                        Reason = "Project access changed during native General configuration. Inspect locally before another operation." });
                }
                return Response.Success(result);
            }
            finally { request.RevalidateProjectPropertyAuthorization = null; }
        }
    }
}
