using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Expose les opérations contrôlées sur les compléments enregistrés du VBE.</summary>
    internal sealed partial class VbeEditorWindows
    {

        /// <summary>Capture l’identité et l’état de connexion lisibles d’un complément.</summary>
        /// <param name="addIn">Complément natif à interroger.</param>
        /// <param name="index">Position facultative dans l’inventaire.</param>
        /// <returns>Un instantané sérialisable et son empreinte si toutes les propriétés ont été lues.</returns>
        private static object AddInSnapshot(dynamic addIn, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "ProgId", () => (string)addIn.ProgId);
            Read(fields, errors, "Guid", () => (string)addIn.Guid);
            Read(fields, errors, "Description", () => (string)addIn.Description);
            Read(fields, errors, "Connect", () => (bool)addIn.Connect);
            string version = null;
            if (errors.Count == 0)
                using (var sha = SHA256.Create())
                    version = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(fields)))).Replace("-", "").ToLowerInvariant();
            return new { Index = index, Properties = fields, Errors = errors, AddInVersion = version };
        }

        /// <summary>Résout un complément enregistré par son ProgID sans accepter les identités ambiguës.</summary>
        /// <param name="progId">ProgID renvoyé par l’inventaire des compléments.</param>
        /// <returns>L’objet natif du complément correspondant.</returns>
        private dynamic FindAddIn(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId) || progId.Length > 255) throw new ArgumentException("ProgId from list_addins is required.");
            dynamic result = null;
            foreach (dynamic item in vbe.AddIns)
                if (string.Equals((string)item.ProgId, progId, StringComparison.OrdinalIgnoreCase))
                {
                    if (result != null) throw new InvalidOperationException("The registered add-in identity is ambiguous.");
                    result = item;
                }
            if (result == null) throw new InvalidOperationException("The registered VBE add-in is no longer present.");
            return result;
        }

        /// <summary>Connecte ou déconnecte un complément après vérification de son état précédemment lu.</summary>
        /// <param name="request">Requête contenant l’action, le ProgID et l’empreinte attendue.</param>
        /// <returns>Le résultat de la mutation, avec état relu et erreurs natives éventuelles.</returns>
        public object SetAddInConnection(Request request)
        {
            if (request.Action != "connect" && request.Action != "disconnect") throw new ArgumentException("Use connect or disconnect.");
            dynamic addIn = FindAddIn(request.ProgId);
            dynamic before = AddInSnapshot(addIn, null);
            if (string.IsNullOrWhiteSpace(request.ExpectedAddInVersion) || before.AddInVersion == null ||
                !string.Equals((string)before.AddInVersion, request.ExpectedAddInVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The add-in state changed or could not be read completely. Read list_addins again.");
            if (string.Equals(request.ProgId, "VBAi.AddIn", StringComparison.OrdinalIgnoreCase) ||
                (Guid.TryParse((string)before.Properties["Guid"], out Guid identity) && identity == typeof(AddIn).GUID))
                throw new InvalidOperationException("VBAi cannot change its own connection from an active agent command.");
            bool desired = request.Action == "connect";
            if ((bool)before.Properties["Connect"] == desired)
                return new { Applied = false, Verified = true, Before = (object)before, After = (object)before, PersistenceVerified = false };
            string nativeError = null;
            try { addIn.Connect = desired; }
            catch (Exception error) { nativeError = error.Message; }
            object after = null;
            bool verified = false;
            string readError = null;
            try
            {
                dynamic current = AddInSnapshot(FindAddIn(request.ProgId), null);
                after = current;
                verified = nativeError == null && current.AddInVersion != null &&
                    string.Equals((string)current.Properties["Guid"], (string)before.Properties["Guid"], StringComparison.OrdinalIgnoreCase) &&
                    (bool)current.Properties["Connect"] == desired;
            }
            catch (Exception error) { readError = error.Message; }
            return new
            {
                Applied = nativeError == null ? (bool?)true : null,
                Verified = verified,
                VerificationPending = !verified,
                Before = (object)before,
                After = after,
                NativeError = nativeError,
                ReadbackError = readError,
                PersistenceVerified = false,
                NextRead = verified ? null : "list_addins"
            };
        }
    }
}
