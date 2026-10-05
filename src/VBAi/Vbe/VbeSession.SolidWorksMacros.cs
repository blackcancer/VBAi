using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{
    internal sealed partial class VbeSession
    {
        private bool macroInFlight, macroQuarantined;
        private int macroAuthorizationDepth;
        // Bridge, chat and editor sessions share one native owning STA.
        [ThreadStatic] private static VbeSession macroOwner;
        [ThreadStatic] private static bool macroOwnerQuarantined;

        private bool MacroDispatchBlocked(string command) => macroQuarantined || macroOwnerQuarantined ||
            ((macroInFlight || macroOwner != null) &&
             !(macroAuthorizationDepth > 0 && (macroOwner == null || ReferenceEquals(macroOwner, this)) && command == "list_projects"));

        private void RequireMacroSettled()
        {
            if (macroInFlight || macroQuarantined || macroOwner != null || macroOwnerQuarantined)
                throw new InvalidOperationException("An original native macro operation is pending or uncertain. Inspect locally; do not retry.");
        }

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
