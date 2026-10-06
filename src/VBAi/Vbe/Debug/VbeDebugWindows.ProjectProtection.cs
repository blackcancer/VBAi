using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;


namespace VBAi
{

    /// <summary>Inspecte et configure le verrouillage natif d’un projet VBA sans exposer ses secrets.</summary>
    internal static partial class VbeDebugWindows
    {

        /// <summary>Sonde injectée du dialogue Protection, ne retournant jamais le texte des secrets.</summary>
        internal interface IProjectProtectionProbe
        {

            /// <summary>Recherche le dialogue exact du projet dans le processus courant.</summary>
            /// <param name="projectName">Nom affiché exact du projet.</param>
            /// <returns>Handle du dialogue unique, ou zéro s’il n’est pas présent.</returns>
            IntPtr Dialog(string projectName);

            /// <summary>Capture le verrouillage, les présences de secrets et l'identité native.</summary>
            /// <param name="dialog">Handle du dialogue natif identifié.</param>
            /// <param name="projectName">Nom de projet utilisé pour revérifier le dialogue.</param>
            /// <returns>État sans contenu des champs secret.</returns>
            ProjectProtectionState Capture(IntPtr dialog, string projectName);

            /// <summary>Écrit la case et les deux secrets sans clavier ni coordonnées.</summary>
            /// <param name="dialog">Handle du dialogue natif.</param>
            /// <param name="projectName">Nom de projet à revalider avant écriture.</param>
            /// <param name="locked">État demandé de la case de verrouillage.</param>
            /// <param name="password">Secret local à écrire dans les champs protégés.</param>
            void Write(IntPtr dialog, string projectName, bool locked, string password);

            /// <summary>Demande la validation native.</summary>
            /// <param name="dialog">Handle du dialogue à valider.</param>
            void Accept(IntPtr dialog);

            /// <summary>Demande l'annulation native.</summary>
            /// <param name="dialog">Handle du dialogue à annuler.</param>
            void Cancel(IntPtr dialog);

            /// <summary>Attend la matérialisation du dialogue.</summary>
            /// <param name="milliseconds">Durée d’attente en millisecondes.</param>
            void Pause(int milliseconds);
        }

        /// <summary>État interne sans contenu de mot de passe.</summary>
        internal sealed class ProjectProtectionState
        {

            /// <summary>Identité native stable du dialogue et de ses contrôles.</summary>
            public string Identity;

            /// <summary>Identité des handles dans une seule ouverture, exclue de la version réouvrable.</summary>
            public string NativeIdentity;

            /// <summary>État observé de la case verrouillage.</summary>
            public bool Locked;

            /// <summary>Longueur observée du premier contrôle secret, jamais son texte.</summary>
            public int PasswordLength;

            /// <summary>Longueur observée de la confirmation, jamais son texte.</summary>
            public int ConfirmationLength;
        }

        /// <summary>Lit les options natives du projet puis annule le dialogue.</summary>
        /// <param name="request">Caption interne portant le nom de projet validé par la file UI.</param>
        /// <returns>État sans secrets, version des options et disponibilité explicite.</returns>
        public static object ReadProjectProtection(Request request) { return ReadProjectProtection(request, new NativeProjectProtectionProbe()); }

        /// <summary>Orchestration injectable de lecture avec annulation garantie autant que possible.</summary>
        /// <param name="request">Sélecteur interne du projet.</param>
        /// <param name="native">Sonde utilisée pour capturer puis annuler le dialogue.</param>
        /// <returns>État sans secret ou indication d’indisponibilité.</returns>
        internal static object ReadProjectProtection(Request request, IProjectProtectionProbe native)
        {
            RequireProtectionRequest(request);
            IntPtr dialog = AwaitProtectionDialog(native, request.Caption);
            if (dialog == IntPtr.Zero) return ProtectionUnavailable();
            bool cancelRequested = false;
            try
            {
                var state = native.Capture(dialog, request.Caption);
                string version = ProtectionVersion(state);
                cancelRequested = true;
                native.Cancel(dialog);
                bool closed = false;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    if (native.Dialog(request.Caption) == IntPtr.Zero) { closed = true; break; }
                    native.Pause(50);
                }
                return new { Available = true, Project = request.Project, ProjectName = request.Caption,
                    LockedForViewing = state.Locked, PasswordPresent = state.PasswordLength > 0,
                    ConfirmationPresent = state.ConfirmationLength > 0,
                    OptionsVersion = version, DialogClosed = closed, PersistenceVerified = false };
            }
            catch { return ProtectionUnavailable(); }
            finally { if (!cancelRequested) TryCancelProtection(native, dialog, request.Caption); }
        }

        /// <summary>Configure le verrouillage du projet depuis un fichier secret local explicitement fourni.</summary>
        /// <param name="request">Action lock/clear, Path secret pour lock et version des options natives.</param>
        /// <returns>Validation demandée et fermeture observée ; sauvegarde et réouverture restent nécessaires.</returns>
        public static object SetProjectProtection(Request request) { return SetProjectProtection(request, new NativeProjectProtectionProbe()); }

        /// <summary>Vérifie toutes les gardes avant écriture et ne transmet aucun secret dans les résultats ou erreurs.</summary>
        /// <param name="request">Action, sélecteur de projet, chemin de secret et version attendue.</param>
        /// <param name="native">Sonde du dialogue utilisée pour capturer, écrire, confirmer ou annuler.</param>
        /// <returns>Résultat vérifié ou état incertain sans divulguer le secret.</returns>
        internal static object SetProjectProtection(Request request, IProjectProtectionProbe native)
        {
            RequireProtectionRequest(request);
            string password = null;
            IntPtr dialog = AwaitProtectionDialog(native, request.Caption);
            if (dialog == IntPtr.Zero) return ProtectionUnavailable();
            bool started = false, commitRequested = false;
            try
            {
                if ((request.Action != "lock" && request.Action != "clear") || string.IsNullOrWhiteSpace(request.ExpectedOptionsVersion))
                    throw new ArgumentException("Action lock/clear and ExpectedOptionsVersion are required.");
                var before = native.Capture(dialog, request.Caption);
                if (!string.Equals(ProtectionVersion(before), request.ExpectedOptionsVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Protection options changed; inspect them again.");
                password = request.Action == "lock" ? ReadProtectionSecret(request.Path) : "";
                started = true;
                bool locked = request.Action == "lock";
                native.Write(dialog, request.Caption, locked, password);
                var after = native.Capture(dialog, request.Caption);
                if (after.Identity != before.Identity || after.NativeIdentity != before.NativeIdentity || after.Locked != locked ||
                    after.PasswordLength != password.Length || after.ConfirmationLength != password.Length)
                    throw new InvalidOperationException("Native protection readback failed.");
                commitRequested = true;
                native.Accept(dialog);
                bool closed = false;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    if (native.Dialog(request.Caption) == IntPtr.Zero) { closed = true; break; }
                    native.Pause(50);
                }
                return new { Available = true, Project = request.Project, ProjectName = request.Caption,
                    CommittedRequested = true, DialogClosed = closed, Closed = closed, ControlValueVerified = true,
                    LockedForViewing = locked, PersistenceVerified = false, RetryAllowed = false,
                    Limit = "Save the host document and reopen it to verify project protection persistence. No save was invoked." };
            }
            catch
            {
                return new { Available = false, Uncertain = started, MutationInvoked = started,
                    CommittedRequested = commitRequested, PersistenceVerified = false, RetryAllowed = false,
                    Reason = started ? "Native protection configuration could not be verified; inspect the project before another operation."
                        : "Protection options are unavailable or changed; inspect them again." };
            }
            finally
            {
                password = null;
                TryCancelProtection(native, dialog, request.Caption);
            }
        }

        /// <summary>Valide les informations internes de sélection sans accepter de nom de dialogue implicite.</summary>
        /// <param name="request">Requête devant porter le sélecteur et le nom de projet validé.</param>
        private static void RequireProtectionRequest(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.Caption))
                throw new ArgumentException("Project and the internally validated project display name are required.");
        }

        /// <summary>Attend brièvement le dialogue exact sans ouvrir un autre projet.</summary>
        /// <param name="native">Sonde qui localise le dialogue dans le processus.</param>
        /// <param name="name">Nom affiché exact du projet.</param>
        /// <returns>Handle observé du dialogue ou zéro à l’expiration.</returns>
        private static IntPtr AwaitProtectionDialog(IProjectProtectionProbe native, string name)
        {
            for (int attempt = 0; attempt < 60; attempt++)
            {
                IntPtr dialog;
                try { dialog = native.Dialog(name); } catch { return IntPtr.Zero; }
                if (dialog != IntPtr.Zero) return dialog;
                native.Pause(50);
            }
            return IntPtr.Zero;
        }

        /// <summary>Retourne une indisponibilité sans détails natifs susceptibles de contenir un secret.</summary>
        /// <returns>Résultat sérialisable indiquant l’absence de contrôles natifs fiables.</returns>
        private static object ProtectionUnavailable()
        {
            return new { Available = false, PersistenceVerified = false,
                Reason = "The exact project Protection dialog or its required native controls are unavailable." };
        }

        /// <summary>Demande Cancel seulement si le dialogue exact reste ouvert, sans exposer les erreurs natives.</summary>
        /// <param name="native">Sonde du dialogue à annuler.</param>
        /// <param name="dialog">Handle de l’ouverture observée.</param>
        /// <param name="name">Nom exact du projet à revérifier.</param>
        private static void TryCancelProtection(IProjectProtectionProbe native, IntPtr dialog, string name)
        {
            try { if (native.Dialog(name) == dialog) native.Cancel(dialog); } catch { }
        }

        /// <summary>Calcule une version à partir de l'identité et des présences de secrets, sans contenu ni longueur.</summary>
        /// <param name="state">État observé sans texte de mot de passe.</param>
        /// <returns>Empreinte SHA-256 des indicateurs non secrets.</returns>
        private static string ProtectionVersion(ProjectProtectionState state)
        {
            string text = state.Identity + "|" + state.Locked + "|" +
                (state.PasswordLength > 0) + "|" + (state.ConfirmationLength > 0);
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant();
        }

        /// <summary>Lit un fichier UTF8 strict borné contenant exactement un secret d'une ligne.</summary>
        /// <param name="path">Chemin absolu du fichier local contenant le secret.</param>
        /// <returns>Secret décodé après vérification de l’encodage et des limites.</returns>
        private static string ReadProtectionSecret(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
                    throw new ArgumentException();
                byte[] bytes;
                using (var stream = OpenProtectionSecret(path))
                {
                    if (stream.Length == 0 || stream.Length > 512) throw new ArgumentException();
                    bytes = new byte[(int)stream.Length];
                    int offset = 0;
                    while (offset < bytes.Length)
                    {
                        int count = stream.Read(bytes, offset, bytes.Length - offset);
                        if (count == 0) throw new ArgumentException();
                        offset += count;
                    }
                }
                string secret = new UTF8Encoding(false, true).GetString(bytes);
                Array.Clear(bytes, 0, bytes.Length);
                // A successful strict decode of the nonempty byte buffer has at least one character.
                if (secret[0] == '\uFEFF') secret = secret.Substring(1);
                if (secret.Length < 1 || secret.Length > 128 || secret.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0)
                    throw new ArgumentException();
                return secret;
            }
            catch { throw new ArgumentException("The local secret file must contain one strict UTF8 line of 1 to 128 characters and at most 512 bytes."); }
        }

        /// <summary>Structure TCITEMW utilisée uniquement dans le processus propriétaire du contrôle.</summary>
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct ProtectionTabItem
        {

            /// <summary>Masque TCIF_TEXT.</summary>
            public uint Mask;

            /// <summary>État natif du tab, inutilisé pour la lecture du texte.</summary>
            public uint State;

            /// <summary>Masque d'état natif, inutilisé.</summary>
            public uint StateMask;

            /// <summary>Adresse locale du tampon Unicode.</summary>
            public IntPtr Text;

            /// <summary>Capacité du tampon en caractères.</summary>
            public int TextCapacity;

            /// <summary>Index d'image, inutilisé.</summary>
            public int Image;

            /// <summary>Donnée associée au tab, inutilisée.</summary>
            public IntPtr Parameter;
        }

        /// <summary>Lit une valeur de style native d’une fenêtre x64.</summary>
        /// <param name="window">Handle de la fenêtre.</param>
        /// <param name="index">Index de style transmis à GetWindowLongPtrW.</param>
        /// <returns>Valeur native du style.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        private static extern IntPtr NativeProtectionWindowLong(IntPtr window, int index);

        /// <summary>Lit le parent natif d’une fenêtre.</summary>
        /// <param name="window">Handle de la fenêtre.</param>
        /// <returns>Handle parent, ou zéro.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetParent")]
        private static extern IntPtr NativeProtectionWindowParent(IntPtr window);

        /// <summary>Frontière injectable de lecture des styles Win32.</summary>
        internal static Func<IntPtr, long> ProtectionWindowStyle = window => NativeProtectionWindowLong(window, -16).ToInt64();

        /// <summary>Frontière injectable de lecture du parent Win32.</summary>
        internal static Func<IntPtr, IntPtr> ProtectionWindowParent = NativeProtectionWindowParent;

        /// <summary>Frontière injectable de lecture des titres tabulaires sans pointeur interprocessus.</summary>
        internal static Func<IntPtr, int, string> ProtectionTabText = ReadProtectionTabText;

        /// <summary>Frontière injectable de changement de page par la feuille de propriétés native.</summary>
        internal static Func<IntPtr, IntPtr, int, bool> ProtectionSelectTab = SelectProtectionTab;

        /// <summary>Opens the bounded local secret read with the production file-sharing policy.</summary>
        internal static Func<string, Stream> OpenProtectionSecret = path => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        /// <summary>Allocates owned native inspection buffers; failures retain the same cleanup path.</summary>
        internal static Func<int, IntPtr> AllocateProtectionBuffer = System.Runtime.InteropServices.Marshal.AllocHGlobal;

        /// <summary>Interdit toute opération sur une fenêtre d'un autre processus.</summary>
        /// <param name="window">Handle dont le processus propriétaire doit être vérifié.</param>
        private static void RequireProtectionOwner(IntPtr window)
        {
            uint pid;
            if (window == IntPtr.Zero || GetWindowThreadProcessId(window, out pid) == 0 ||
                pid != (uint)Process.GetCurrentProcess().Id)
                throw new InvalidOperationException("The protection window does not belong to this host process.");
        }

        /// <summary>Lit TCITEMW avec buffers locaux après vérification du PID, sans accès à un Edit.</summary>
        /// <param name="tab">Handle du contrôle d’onglets natif.</param>
        /// <param name="index">Index de l’onglet à lire.</param>
        /// <returns>Libellé Unicode exact de l’onglet.</returns>
        private static string ReadProtectionTabText(IntPtr tab, int index)
        {
            RequireProtectionOwner(tab);
            IntPtr text = IntPtr.Zero, item = IntPtr.Zero;
            try
            {
                text = AllocateProtectionBuffer(512);
                for (int offset = 0; offset < 512; offset += 2)
                    System.Runtime.InteropServices.Marshal.WriteInt16(text, offset, 0);
                var descriptor = new ProtectionTabItem { Mask = 1, Text = text, TextCapacity = 256 };
                item = AllocateProtectionBuffer(System.Runtime.InteropServices.Marshal.SizeOf(typeof(ProtectionTabItem)));
                System.Runtime.InteropServices.Marshal.StructureToPtr(descriptor, item, false);
                if (SendMessageInt(tab, 0x133C, new IntPtr(index), item) == IntPtr.Zero)
                    throw new InvalidOperationException("The native protection tab title is unreadable.");
                var observed = (ProtectionTabItem)System.Runtime.InteropServices.Marshal.PtrToStructure(item, typeof(ProtectionTabItem));
                string caption = observed.Text == IntPtr.Zero ? "" : System.Runtime.InteropServices.Marshal.PtrToStringUni(observed.Text);
                if (caption.Length > 255) throw new InvalidOperationException("The native tab title exceeds the bounded inspection limit.");
                return caption;
            }
            finally
            {
                if (item != IntPtr.Zero) System.Runtime.InteropServices.Marshal.FreeHGlobal(item);
                if (text != IntPtr.Zero) System.Runtime.InteropServices.Marshal.FreeHGlobal(text);
            }
        }

        /// <summary>PSM_SETCURSEL active la page et déclenche ses notifications natives PSN_KILLACTIVE/PSN_SETACTIVE.</summary>
        /// <param name="dialog">Handle de la feuille de propriétés.</param>
        /// <param name="tab">Handle de son contrôle d’onglets.</param>
        /// <param name="index">Index de page à activer.</param>
        /// <returns><see langword="true"/> si la sélection native a été confirmée.</returns>
        private static bool SelectProtectionTab(IntPtr dialog, IntPtr tab, int index)
        {
            RequireProtectionOwner(dialog); RequireProtectionOwner(tab);
            return SendMessageInt(dialog, 0x0465, new IntPtr(index), IntPtr.Zero) != IntPtr.Zero &&
                SendMessageInt(tab, 0x130B, IntPtr.Zero, IntPtr.Zero).ToInt32() == index;
        }

        /// <summary>Sonde Win32 des seuls contrôles exacts du dialogue de propriétés du projet courant.</summary>
        private sealed class NativeProjectProtectionProbe : IProjectProtectionProbe
        {

            /// <summary>Nom de projet utilisé pour revalider le dialogue avant tout clic natif.</summary>
            private string validatedProjectName;

            /// <summary>Recherche exactement un dialogue du projet dans le processus hôte.</summary>
            /// <param name="projectName">Nom du projet utilisé dans le titre du dialogue.</param>
            /// <returns>Handle unique du dialogue ou zéro s’il est absent.</returns>
            public IntPtr Dialog(string projectName)
            {
                var matches = new List<IntPtr>();
                uint owner = (uint)Process.GetCurrentProcess().Id;
                string[] titles = { projectName + " - Project Properties", projectName + " - Propriétés du projet" };
                EnumWindows((window, data) =>
                {
                    uint pid; GetWindowThreadProcessId(window, out pid);
                    if (pid == owner && ClassName(window) == "#32770" && IsWindowVisible(window) &&
                        titles.Contains(WindowText(window), StringComparer.Ordinal)) matches.Add(window);
                    return true;
                }, IntPtr.Zero);
                if (matches.Count > 1) throw new InvalidOperationException("Ambiguous project properties dialog.");
                validatedProjectName = projectName;
                return matches.Count == 1 ? matches[0] : IntPtr.Zero;
            }

            /// <summary>Sélectionne l'onglet Protection et capture uniquement les longueurs et le verrouillage.</summary>
            /// <param name="dialog">Handle du dialogue identifié.</param>
            /// <param name="projectName">Nom du projet à vérifier avant lecture.</param>
            /// <returns>Identités, état de verrouillage et longueurs sans valeur de secret.</returns>
            public ProjectProtectionState Capture(IntPtr dialog, string projectName)
            {
                var controls = ProtectionControls(dialog, projectName);
                int checkedState = ReadProtectionCheck(controls[0]);
                return new ProjectProtectionState {
                    Identity = projectName + "|" + WindowText(dialog) + "|" + string.Join("|", controls.Select(item => GetDlgCtrlID(item).ToString())),
                    NativeIdentity = dialog.ToInt64() + "|" + string.Join("|", controls.Select(item => item.ToInt64().ToString())),
                    Locked = checkedState == 1,
                    PasswordLength = PasswordLength(controls[1]), ConfirmationLength = PasswordLength(controls[2]) };
            }

            /// <summary>Configure la case par BM_CLICK et les deux secrets par EM_REPLACESEL, sans clavier.</summary>
            /// <param name="dialog">Handle du dialogue natif.</param>
            /// <param name="projectName">Nom du projet à revérifier.</param>
            /// <param name="locked">État demandé de verrouillage pour affichage.</param>
            /// <param name="password">Secret à écrire dans les champs Password et Confirmation.</param>
            public void Write(IntPtr dialog, string projectName, bool locked, string password)
            {
                var controls = ProtectionControls(dialog, projectName);
                if (locked && ReadProtectionCheck(controls[0]) != 1)
                {
                    SendMessageInt(controls[0], BmClick, IntPtr.Zero, IntPtr.Zero);
                    if (ReadProtectionCheck(controls[0]) != 1) throw new InvalidOperationException();
                }
                controls = ProtectionControls(dialog, projectName);
                SetPassword(controls[1], password);
                SetPassword(controls[2], password);
                if (!locked && ReadProtectionCheck(controls[0]) != 0)
                {
                    SendMessageInt(controls[0], BmClick, IntPtr.Zero, IntPtr.Zero);
                    if (ReadProtectionCheck(controls[0]) != 0) throw new InvalidOperationException();
                }
            }

            /// <summary>Demande OK par le bouton natif visible et activé.</summary>
            /// <param name="dialog">Handle du dialogue natif.</param>
            public void Accept(IntPtr dialog) { ProtectionButton(dialog, 1); }

            /// <summary>Demande Cancel par le bouton natif visible et activé.</summary>
            /// <param name="dialog">Handle du dialogue natif.</param>
            public void Cancel(IntPtr dialog) { ProtectionButton(dialog, 2); }

            /// <summary>Attend le traitement des messages natifs.</summary>
            /// <param name="milliseconds">Durée d’attente bornée en millisecondes.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }

            /// <summary>Exige un titre de tab exact et les contrôles observés de la seule page Protection visible.</summary>
            /// <param name="dialog">Handle de la feuille de propriétés.</param>
            /// <param name="projectName">Nom exact du projet à revalider.</param>
            /// <returns>Case de verrouillage puis champs mot de passe et confirmation.</returns>
            private IntPtr[] ProtectionControls(IntPtr dialog, string projectName)
            {
                if (Dialog(projectName) != dialog) throw new InvalidOperationException();
                RequireProtectionOwner(dialog);
                var descendants = ProtectionChildren(dialog);
                var tabs = descendants.Where(window => ClassName(window) == "SysTabControl32" &&
                    IsWindowVisible(window) && OptionsWindowEnabled(window) &&
                    ProtectionWindowParent(window) == dialog).ToArray();
                if (tabs.Length != 1) throw new InvalidOperationException();
                IntPtr tab = tabs[0]; RequireProtectionOwner(tab);
                int count = SendMessageInt(tab, 0x1304, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (count < 1 || count > 16) throw new InvalidOperationException();
                int selected = -1;
                for (int index = 0; index < count; index++)
                    if (ProtectionTabText(tab, index) == "Protection")
                    {
                        if (selected >= 0) throw new InvalidOperationException();
                        selected = index;
                    }
                if (selected < 0) throw new InvalidOperationException();
                if (SendMessageInt(tab, 0x130B, IntPtr.Zero, IntPtr.Zero).ToInt32() != selected &&
                    !ProtectionSelectTab(dialog, tab, selected)) throw new InvalidOperationException();
                if (SendMessageInt(tab, 0x130B, IntPtr.Zero, IntPtr.Zero).ToInt32() != selected)
                    throw new InvalidOperationException();
                descendants = ProtectionChildren(dialog);
                string[] captions = { "Lock project for viewing", "Verrouiller le projet pour l'affichage",
                    "Verrouiller le projet pour l’affichage" };
                var checks = descendants.Where(window => ClassName(window) == "Button" &&
                    IsWindowVisible(window) && OptionsWindowEnabled(window) &&
                    captions.Contains(WindowText(window).Replace("&", ""), StringComparer.Ordinal)).ToArray();
                if (checks.Length != 1) throw new InvalidOperationException();
                IntPtr checkbox = checks[0]; RequireProtectionOwner(checkbox);
                long checkStyle = ProtectionWindowStyle(checkbox) & 0xf;
                if ((checkStyle != 2 && checkStyle != 3) || GetDlgCtrlID(checkbox) != 5463)
                    throw new InvalidOperationException();
                IntPtr page = ProtectionWindowParent(checkbox); RequireProtectionOwner(page);
                if (ClassName(page) != "#32770" || !IsWindowVisible(page) ||
                    ProtectionWindowParent(page) != dialog) throw new InvalidOperationException();
                var passwords = descendants.Where(window => ClassName(window) == "Edit" &&
                    IsWindowVisible(window) && ProtectionWindowParent(window) == page &&
                    (ProtectionWindowStyle(window) & 0x20) != 0).ToArray();
                if (passwords.Length != 2 || passwords[0] == passwords[1]) throw new InvalidOperationException();
                foreach (IntPtr password in passwords) RequireProtectionOwner(password);
                passwords = passwords.OrderBy(window => GetDlgCtrlID(window)).ToArray();
                if (GetDlgCtrlID(passwords[0]) != 5461 || GetDlgCtrlID(passwords[1]) != 5462)
                    throw new InvalidOperationException();
                return new[] { checkbox, passwords[0], passwords[1] };
            }

            /// <summary>Énumère les vrais enfants Win32 du dialogue, distincts des volets VbaWindow.</summary>
            /// <param name="dialog">Handle du dialogue parent.</param>
            /// <returns>Handles de tous ses enfants, dans l’ordre d’énumération Windows.</returns>
            private static List<IntPtr> ProtectionChildren(IntPtr dialog)
            {
                var result = new List<IntPtr>();
                EnumChildWindows(dialog, (window, ignored) => {
                    if (result.Count >= 256) throw new InvalidOperationException("The native protection dialog has too many child windows.");
                    result.Add(window); return true;
                }, IntPtr.Zero);
                return result;
            }

            /// <summary>Lit BM_GETCHECK et refuse l'état indéterminé ou une réponse inconnue.</summary>
            /// <param name="checkbox">Case native qualifiée.</param>
            /// <returns>État BM_GETCHECK, zéro ou un.</returns>
            private static int ReadProtectionCheck(IntPtr checkbox)
            {
                RequireProtectionOwner(checkbox);
                int state = SendMessageInt(checkbox, 0x00F0, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (state != 0 && state != 1) throw new InvalidOperationException();
                return state;
            }

            /// <summary>Lit seulement WM_GETTEXTLENGTH, jamais WM_GETTEXT sur un Edit.</summary>
            /// <param name="password">Champ mot de passe natif.</param>
            /// <returns>Longueur du texte sans lire son contenu.</returns>
            private static int PasswordLength(IntPtr password)
            {
                RequireProtectionOwner(password);
                int length = SendMessageInt(password, 0x000E, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (length < 0 || length > 128) throw new InvalidOperationException();
                return length;
            }

            /// <summary>Remplace le secret dans un Edit ES_PASSWORD natif et relit exclusivement sa longueur.</summary>
            /// <param name="window">Handle du champ secret.</param>
            /// <param name="password">Secret à écrire, ou chaîne vide.</param>
            private static void SetPassword(IntPtr window, string password)
            {
                RequireProtectionOwner(window);
                if (ClassName(window) != "Edit" || (ProtectionWindowStyle(window) & 0x20) == 0)
                    throw new InvalidOperationException();
                if (!OptionsWindowEnabled(window))
                {
                    if (password.Length == 0 && PasswordLength(window) == 0) return;
                    throw new InvalidOperationException();
                }
                SendMessageInt(window, 0x00B1, IntPtr.Zero, new IntPtr(-1));
                SendMessageText(window, 0x00C2, IntPtr.Zero, password);
                if (PasswordLength(window) != password.Length) throw new InvalidOperationException();
            }

            /// <summary>Valide le projet, le PID et l'identité native du bouton avant BM_CLICK.</summary>
            /// <param name="dialog">Handle du dialogue du projet actif.</param>
            /// <param name="id">Identifiant natif du bouton OK ou Cancel.</param>
            private void ProtectionButton(IntPtr dialog, int id)
            {
                if (validatedProjectName == null || Dialog(validatedProjectName) != dialog)
                    throw new InvalidOperationException();
                RequireProtectionOwner(dialog);
                IntPtr button = GetDlgItem(dialog, id); RequireProtectionOwner(button);
                if (ClassName(button) != "Button" || ProtectionWindowParent(button) != dialog ||
                    GetDlgCtrlID(button) != id || !IsWindowVisible(button) || !OptionsWindowEnabled(button) ||
                    !PostMessage(button, BmClick, IntPtr.Zero, IntPtr.Zero)) throw new InvalidOperationException();
            }
        }
    }
}
