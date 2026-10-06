using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Serializes SOLIDWORKS macro creation/publication on the VBE STA and quarantines the session after uncertain native effects.</summary>
    internal sealed partial class VbeSession
    {

        /// <summary>Tracks whether this session has an active native macro call or a non-retryable uncertain outcome.</summary>
        private bool macroInFlight, macroQuarantined;

        /// <summary>Allows only the nested project inventory reads required while revalidating native macro authorization.</summary>
        private int macroAuthorizationDepth;
        // Bridge, chat and editor sessions share one native owning STA.
        /// <summary>Thread-local session currently owning a native macro operation on this shared VBE STA.</summary>
        [ThreadStatic] private static VbeSession macroOwner;

        /// <summary>Thread-local quarantine that blocks other sessions on the same STA after an uncertain macro outcome.</summary>
        [ThreadStatic] private static bool macroOwnerQuarantined;

        /// <summary>Blocks dispatch during or after native macro ownership, allowing only the scoped list-projects authorization read.</summary>
        /// <param name="command">Command being considered for dispatch.</param>
        /// <returns><see langword="true"/> when this command cannot enter the shared native VBE route.</returns>
        private bool MacroDispatchBlocked(string command) => macroQuarantined || macroOwnerQuarantined ||
            ((macroInFlight || macroOwner != null) &&
             !(macroAuthorizationDepth > 0 && (macroOwner == null || ReferenceEquals(macroOwner, this)) && command == "list_projects"));

        /// <summary>Throws while a native macro operation is active or quarantined, preventing a retry after uncertainty.</summary>
        private void RequireMacroSettled()
        {
            if (macroInFlight || macroQuarantined || macroOwner != null || macroOwnerQuarantined)
                throw new InvalidOperationException("An original native macro operation is pending or uncertain. Inspect locally; do not retry.");
        }

        /// <summary>Runs one authorized create or publish operation on the owning VBE STA and quarantines the route after an uncertain result.</summary>
        /// <param name="request">Exact create/publish request whose command, project, path, revision, and mode are frozen during revalidation.</param>
        /// <returns>Native macro operation result with terminal and uncertainty evidence.</returns>
        internal async Task<object> SolidWorksMacroAsync(Request request)
        {
            RequireGeneralSettled();
            if (request == null || (request.Command != "create_solidworks_macro" && request.Command != "publish_solidworks_macro"))
                throw new ArgumentException("An explicit native macro operation is required.");
            if (bridgeOperationsInFlight != 0)
                throw new InvalidOperationException("A bridge operation must settle before native macro creation/publication.");
            int owner = Thread.CurrentThread.ManagedThreadId;
            Action context = () => {
                if (Thread.CurrentThread.ManagedThreadId != owner || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Native macro operation left its owning VBE STA.");
            };
            context();
            var authorization = request.RevalidateMacroAuthorization;
            if (authorization == null) throw new InvalidOperationException("Original native macro authorization is required.");
            string command = request.Command, project = request.Project, path = request.Path, version = request.ExpectedProjectVersion;
            int mode = request.ExpectedMode;
            bool claimed = false;
            Action<bool> scoped = live => {
                context();
                if (request.Command != command || request.Project != project || request.Path != path ||
                    request.ExpectedProjectVersion != version || request.ExpectedMode != mode)
                    throw new InvalidOperationException("Original native macro request changed.");
                if (live) macroAuthorizationDepth++;
                try { authorization(live); }
                finally { if (live) macroAuthorizationDepth--; }
            };
            Action<VbeProjectComponents.MacroMutationClaim> journal = claim => {
                context();
                // No source, project identifiers or user paths enter the global log.
                LoadLog.AppendText(LoadLog.PathName, DateTime.UtcNow.ToString("o") + " Native macro claim Phase=" +
                    (claim.Phase ?? "Unknown").Split(':')[0] + " Ordinal=" + claim.Ordinal + Environment.NewLine);
                if (claim.Phase != "Terminal") claimed = true;
            };
            Func<Request, Task<object>> priorGeneralReader = components.PublicationReadGeneral;
            Func<Request, Task<object>> ownedGeneralReader = read => ReadPublicationGeneralAsync(read, scoped, context);
            if (command == "publish_solidworks_macro") components.PublicationReadGeneral = ownedGeneralReader;
            macroInFlight = true;
            macroOwner = this;
            try
            {
                object result = command == "create_solidworks_macro"
                    ? (object)await components.CreateSolidWorksMacroAsync(request, scoped, journal, context)
                    : await components.PublishSolidWorksMacroAsync(request, scoped, journal, context);
                var serializer = new JavaScriptSerializer();
                var fields = serializer.DeserializeObject(serializer.Serialize(result)) as IDictionary<string, object>;
                bool terminal = fields != null && fields.TryGetValue("Terminal", out object t) && t is bool && (bool)t;
                bool uncertain = fields == null || !fields.TryGetValue("Uncertain", out object u) || !(u is bool) || (bool)u;
                macroQuarantined = uncertain || !terminal;
                return result;
            }
            catch
            {
                if (claimed) macroQuarantined = true;
                throw;
            }
            finally
            {
                if (command == "publish_solidworks_macro" && ReferenceEquals(components.PublicationReadGeneral, ownedGeneralReader))
                    components.PublicationReadGeneral = priorGeneralReader;
                macroOwnerQuarantined |= macroQuarantined;
                macroOwner = null;
                macroInFlight = false;
            }
        }
    }
}
