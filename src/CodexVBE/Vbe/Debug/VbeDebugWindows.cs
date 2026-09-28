using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace CodexVBE
{
    /// <summary>Observe et pilote les fenêtres natives de débogage et de dialogue du VBE via Win32, UI Automation et MSAA.</summary>
    // Run these accessibility calls away from the VBE UI thread. UIA and MSAA
    // query native windows, not the VBIDE object model used by VbeDebug.
    internal static partial class VbeDebugWindows
    {
        /// <summary>Callback de l’énumération Win32 des fenêtres.</summary>
        /// <param name="handle">Handle de la fenêtre énumérée.</param>
        /// <param name="parameter">Paramètre transmis à l’énumération.</param>
        /// <returns><see langword="true"/> pour poursuivre l’énumération.</returns>
        internal delegate bool EnumWindowCallback(IntPtr handle, IntPtr parameter);

        /// <summary>Énumère les fenêtres de premier niveau.</summary>
        /// <param name="callback">Callback invoqué pour chaque fenêtre.</param>
        /// <param name="parameter">Paramètre transmis au callback.</param>
        /// <returns><see langword="true"/> si l’énumération a réussi.</returns>
        [DllImport("user32.dll", EntryPoint = "EnumWindows")] private static extern bool NativeEnumWindows(EnumWindowCallback callback, IntPtr parameter);
        /// <summary>Énumère les fenêtres enfants d’une fenêtre parente.</summary>
        /// <param name="parent">Fenêtre dont les enfants sont énumérés.</param>
        /// <param name="callback">Callback invoqué pour chaque enfant.</param>
        /// <param name="parameter">Paramètre transmis au callback.</param>
        /// <returns><see langword="true"/> si l’énumération a réussi.</returns>
        [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool NativeEnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        /// <summary>Obtient le PID et le thread propriétaires d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <param name="processId">Reçoit le PID propriétaire.</param>
        /// <returns>Identifiant du thread propriétaire, ou zéro si absent.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint NativeGetWindowThreadProcessId(IntPtr handle, out uint processId);
        /// <summary>Lit la classe Win32 d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <param name="text">Tampon recevant le nom de classe.</param>
        /// <param name="capacity">Capacité du tampon.</param>
        /// <returns>Nombre de caractères copiés.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassName")] private static extern int NativeGetClassName(IntPtr handle, StringBuilder text, int capacity);
        /// <summary>Lit le titre Win32 d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <param name="text">Tampon recevant le titre.</param>
        /// <param name="capacity">Capacité du tampon.</param>
        /// <returns>Nombre de caractères copiés.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowText")] private static extern int NativeGetWindowText(IntPtr handle, StringBuilder text, int capacity);
        /// <summary>Obtient l’identifiant de contrôle associé à un handle de dialogue.</summary>
        /// <param name="handle">Handle du contrôle.</param>
        /// <returns>Identifiant numérique du contrôle.</returns>
        [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int NativeGetDlgCtrlID(IntPtr handle);
        /// <summary>Recherche un contrôle enfant de dialogue par identifiant.</summary>
        /// <param name="handle">Handle du dialogue.</param>
        /// <param name="controlId">Identifiant du contrôle.</param>
        /// <returns>Handle du contrôle trouvé, ou zéro.</returns>
        [DllImport("user32.dll", EntryPoint = "GetDlgItem")] private static extern IntPtr NativeGetDlgItem(IntPtr handle, int controlId);
        /// <summary>Indique si une fenêtre est visible.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <returns><see langword="true"/> si la fenêtre est visible.</returns>
        [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool NativeIsWindowVisible(IntPtr handle);
        /// <summary>Poste un message asynchrone dans la file d’une fenêtre.</summary>
        /// <param name="handle">Fenêtre destinataire.</param>
        /// <param name="message">Identifiant du message.</param>
        /// <param name="wParam">Premier paramètre du message.</param>
        /// <param name="lParam">Second paramètre du message.</param>
        /// <returns><see langword="true"/> si le message a été posté.</returns>
        [DllImport("user32.dll", EntryPoint = "PostMessage")] private static extern bool NativePostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        /// <summary>Envoie un message Win32 dont le paramètre lParam est du texte Unicode.</summary>
        /// <param name="handle">Fenêtre destinataire.</param>
        /// <param name="message">Identifiant du message.</param>
        /// <param name="wParam">Premier paramètre du message.</param>
        /// <param name="text">Texte transmis dans lParam.</param>
        /// <returns>Valeur renvoyée par la procédure de fenêtre.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr NativeSendMessageText(IntPtr handle, int message, IntPtr wParam, string text);
        /// <summary>Envoie un message Win32 avec des paramètres entiers.</summary>
        /// <param name="handle">Fenêtre destinataire.</param>
        /// <param name="message">Identifiant du message.</param>
        /// <param name="wParam">Premier paramètre du message.</param>
        /// <param name="lParam">Second paramètre du message.</param>
        /// <returns>Valeur renvoyée par la procédure de fenêtre.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr NativeSendMessageInt(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        /// <summary>Obtient l’interface MSAA associée à la partie cliente d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <param name="objectId">Identifiant MSAA de l’objet visé.</param>
        /// <param name="interfaceId">GUID de l’interface demandée.</param>
        /// <param name="accessible">Reçoit l’objet d’accessibilité.</param>
        /// <returns>Code HRESULT de l’appel.</returns>
        [DllImport("oleacc.dll", EntryPoint = "AccessibleObjectFromWindow")] private static extern int NativeAccessibleObjectFromWindow(IntPtr handle, uint objectId,
            ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object accessible);

        /// <summary>Fonction injectable qui retourne le processus propriétaire d’une fenêtre native.</summary>
        /// <param name="handle">Handle de la fenêtre.</param>
        /// <param name="processId">Reçoit le PID propriétaire.</param>
        /// <returns>Identifiant du thread propriétaire.</returns>
        internal delegate uint WindowProcessReader(IntPtr handle, out uint processId);
        /// <summary>Fonction injectable qui obtient une interface d’accessibilité MSAA.</summary>
        /// <param name="handle">Handle de la fenêtre.</param>
        /// <param name="objectId">Identifiant d’objet MSAA.</param>
        /// <param name="interfaceId">GUID de l’interface demandée.</param>
        /// <param name="accessible">Reçoit l’objet d’accessibilité.</param>
        /// <returns>Code HRESULT de l’appel natif.</returns>
        internal delegate int AccessibleClientReader(IntPtr handle, uint objectId, ref Guid interfaceId, out object accessible);

        /// <summary>Délégué injectable pour l’énumération des fenêtres de premier niveau.</summary>
        internal static Func<EnumWindowCallback, IntPtr, bool> EnumWindows = NativeEnumWindows;
        /// <summary>Délégué injectable pour l’énumération des fenêtres enfants.</summary>
        internal static Func<IntPtr, EnumWindowCallback, IntPtr, bool> EnumChildWindows = NativeEnumChildWindows;
        /// <summary>Délégué injectable pour lire le processus propriétaire d’une fenêtre.</summary>
        internal static WindowProcessReader GetWindowThreadProcessId = NativeGetWindowThreadProcessId;
        /// <summary>Délégué injectable pour lire le nom de classe natif.</summary>
        internal static Func<IntPtr, StringBuilder, int, int> GetClassName = NativeGetClassName;
        /// <summary>Délégué injectable pour lire le texte d’une fenêtre.</summary>
        internal static Func<IntPtr, StringBuilder, int, int> GetWindowText = NativeGetWindowText;
        /// <summary>Délégué injectable pour lire l’identifiant d’un contrôle.</summary>
        internal static Func<IntPtr, int> GetDlgCtrlID = NativeGetDlgCtrlID;
        /// <summary>Délégué injectable pour obtenir un contrôle enfant par identifiant.</summary>
        internal static Func<IntPtr, int, IntPtr> GetDlgItem = NativeGetDlgItem;
        /// <summary>Délégué injectable pour tester la visibilité d’une fenêtre.</summary>
        internal static Func<IntPtr, bool> IsWindowVisible = NativeIsWindowVisible;
        /// <summary>Délégué injectable pour poster un message natif.</summary>
        internal static Func<IntPtr, int, IntPtr, IntPtr, bool> PostMessage = NativePostMessage;
        /// <summary>Délégué injectable pour envoyer un message Unicode.</summary>
        internal static Func<IntPtr, int, IntPtr, string, IntPtr> SendMessageText = NativeSendMessageText;
        /// <summary>Délégué injectable pour envoyer un message à paramètres entiers.</summary>
        internal static Func<IntPtr, int, IntPtr, IntPtr, IntPtr> SendMessageInt = NativeSendMessageInt;
        /// <summary>Délégué injectable pour acquérir un objet d’accessibilité.</summary>
        internal static AccessibleClientReader AccessibleObjectFromWindow = NativeAccessibleObjectFromWindow;
        /// <summary>Action de temporisation injectable entre les interactions native asynchrones.</summary>
        internal static Action<int> PauseNative = Thread.Sleep;

        /// <summary>Identifiant du message standard d’activation d’un bouton.</summary>
        private const int BmClick = 0x00F5;
        /// <summary>Identifiant du message standard de saisie d’un caractère.</summary>
        private const int WmChar = 0x0102;
        /// <summary>Identifiant du message standard d’appui sur une touche.</summary>
        private const int WmKeyDown = 0x0100;
        /// <summary>Identifiant du message standard de relâchement d’une touche.</summary>
        private const int WmKeyUp = 0x0101;
        /// <summary>Code virtuel de la touche Entrée.</summary>
        private const int VkReturn = 13;
        /// <summary>Identifiant MSAA de l’objet client d’une fenêtre.</summary>
        private const uint ObjidClient = 4294967292;
        /// <summary>IID de l’interface standard IAccessible.</summary>
        private static readonly Guid IidAccessible = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");

        /// <summary>Instantané des informations natives utiles sur un contrôle de dialogue.</summary>
        internal sealed class NativeControl
        {
            /// <summary>Handle Win32 du contrôle.</summary>
            public IntPtr Handle;
            /// <summary>Classe Win32 du contrôle.</summary>
            public string Kind;
            /// <summary>Texte actuellement affiché par le contrôle.</summary>
            public string Text;
            /// <summary>Indique si le contrôle est visible.</summary>
            public bool Visible;
        }

        /// <summary>Abstraction des lectures et interactions avec les fenêtres natives du VBE.</summary>
        internal interface INativeProbe
        {
            /// <summary>Retourne la fenêtre racine du VBE.</summary>
            /// <returns>Handle racine, ou zéro si absent.</returns>
            IntPtr VbeRoot();
            /// <summary>Énumère les fenêtres enfants d’un handle.</summary>
            /// <param name="root">Handle parent.</param>
            /// <returns>Handles des fenêtres enfants.</returns>
            List<IntPtr> Children(IntPtr root);
            /// <summary>Recherche un volet parmi les fenêtres énumérées.</summary>
            /// <param name="panes">Fenêtres candidates.</param>
            /// <param name="names">Noms possibles du volet.</param>
            /// <returns>Handle du volet trouvé ou zéro.</returns>
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            /// <summary>Lit une liste accessible dans un volet de débogage.</summary>
            /// <param name="handle">Handle du volet.</param>
            /// <returns>Contenu lu.</returns>
            object List(IntPtr handle);
            /// <summary>Lit le contenu de la fenêtre Immediate.</summary>
            /// <param name="handle">Handle du volet Immediate.</param>
            /// <returns>Instantané de son contenu.</returns>
            object Immediate(IntPtr handle);
            /// <summary>Lit la pile d’appels liée au volet Locals.</summary>
            /// <param name="locals">Handle du volet Locals.</param>
            /// <returns>Trames lues dans la pile d’appels.</returns>
            object CallStack(IntPtr locals);
            /// <summary>Recherche un dialogue dont le titre correspond aux valeurs fournies.</summary>
            /// <param name="titles">Titres possibles.</param>
            /// <returns>Handle du dialogue ou zéro.</returns>
            IntPtr Dialog(params string[] titles);
            /// <summary>Énumère les contrôles Win32 d’un dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Contrôles trouvés.</returns>
            List<NativeControl> DialogControls(IntPtr dialog);
            /// <summary>Lit le message accessible d’un dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Message lu.</returns>
            string DialogMessage(IntPtr dialog);
            /// <summary>Envoie un clic au contrôle.</summary>
            /// <param name="handle">Handle du contrôle.</param>
            /// <returns>Indique si le clic a été transmis.</returns>
            bool Click(IntPtr handle);
            /// <summary>Vérifie la visibilité du contrôle.</summary>
            /// <param name="handle">Handle du contrôle.</param>
            /// <returns><see langword="true"/> si visible.</returns>
            bool Visible(IntPtr handle);
            /// <summary>Attend le nombre de millisecondes indiqué.</summary>
            /// <param name="milliseconds">Durée d’attente.</param>
            void Pause(int milliseconds);
        }

        /// <summary>Abstraction native injectable des boîtes Add Watch, Edit Watch et Quick Watch.</summary>
        internal interface IWatchProbe
        {
            /// <summary>Recherche un dialogue selon ses titres possibles.</summary>
            /// <param name="titles">Titres possibles.</param>
            /// <returns>Handle du dialogue ou zéro.</returns>
            IntPtr Dialog(params string[] titles);
            /// <summary>Recherche un contrôle de dialogue par son identifiant.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <param name="id">Identifiant du contrôle.</param>
            /// <returns>Handle du contrôle ou zéro.</returns>
            IntPtr Item(IntPtr dialog, int id);
            /// <summary>Lit le texte d’un contrôle.</summary>
            /// <param name="handle">Handle du contrôle.</param>
            /// <returns>Texte du contrôle.</returns>
            string Text(IntPtr handle);
            /// <summary>Active un contrôle de dialogue.</summary>
            /// <param name="handle">Handle du contrôle à activer.</param>
            /// <returns><see langword="true"/> si le clic a été posté.</returns>
            bool Click(IntPtr handle);
            /// <summary>Lit l’état coché d’un contrôle bouton.</summary>
            /// <param name="handle">Handle du bouton.</param>
            /// <returns><see langword="true"/> si le bouton est coché.</returns>
            bool Checked(IntPtr handle);
            /// <summary>Remplace le texte d’un contrôle.</summary>
            /// <param name="handle">Handle du champ texte.</param>
            /// <param name="text">Valeur à saisir.</param>
            void Replace(IntPtr handle, string text);
            /// <summary>Attend la durée indiquée.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            void Pause(int milliseconds);
            /// <summary>Lit le message accessible d’une boîte de dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Message accessible.</returns>
            string Message(IntPtr dialog);
            /// <summary>Ferme le dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            void Close(IntPtr dialog);
            /// <summary>Retourne la fenêtre racine du VBE.</summary>
            /// <returns>Handle racine, ou zéro si absent.</returns>
            IntPtr VbeRoot();
            /// <summary>Énumère les fenêtres enfants.</summary>
            /// <param name="root">Handle parent.</param>
            /// <returns>Handles enfants.</returns>
            List<IntPtr> Children(IntPtr root);
            /// <summary>Recherche un volet nommé.</summary>
            /// <param name="panes">Fenêtres candidates.</param>
            /// <param name="names">Noms possibles.</param>
            /// <returns>Handle du volet ou zéro.</returns>
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            /// <summary>Lit le contenu d’une liste du volet Watches.</summary>
            /// <param name="handle">Handle de liste.</param>
            /// <returns>Instantané de liste.</returns>
            object List(IntPtr handle);
            /// <summary>Compte les lignes de surveillance correspondant à l’expression et au contexte.</summary>
            /// <param name="pane">Volet Watches.</param>
            /// <param name="expression">Expression recherchée.</param>
            /// <param name="context">Contexte facultatif.</param>
            /// <returns>Nombre de lignes correspondantes.</returns>
            int WatchMatches(IntPtr pane, string expression, string context);
            /// <summary>Sélectionne une ligne de surveillance correspondante.</summary>
            /// <param name="pane">Volet Watches.</param>
            /// <param name="expression">Expression recherchée.</param>
            /// <param name="context">Contexte facultatif.</param>
            /// <returns><see langword="true"/> si la ligne a été sélectionnée.</returns>
            bool SelectWatchRow(IntPtr pane, string expression, string context);
        }

        /// <summary>Implémente les interactions avec les dialogues de surveillance par Win32 et UI Automation.</summary>
        private sealed class NativeWatchProbe : IWatchProbe
        {
            /// <summary>Lignes UI Automation issues de la dernière recherche des surveillances.</summary>
            private List<AutomationElement> lastRows;
            /// <summary>Recherche un dialogue de surveillance par titre.</summary>
            /// <param name="titles">Titres possibles du dialogue.</param>
            /// <returns>Handle du dialogue, ou zéro si absent.</returns>
            public IntPtr Dialog(params string[] titles) { return FindDialog(titles); }
            /// <summary>Retourne le handle du contrôle de dialogue demandé.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <param name="id">Identifiant numérique du contrôle.</param>
            /// <returns>Handle du contrôle, ou zéro si absent.</returns>
            public IntPtr Item(IntPtr dialog, int id) { return GetDlgItem(dialog, id); }
            /// <summary>Lit le texte visible d’une fenêtre.</summary>
            /// <param name="handle">Handle du contrôle.</param>
            /// <returns>Texte observé.</returns>
            public string Text(IntPtr handle) { return WindowText(handle); }
            /// <summary>Poste un clic standard sur le contrôle.</summary>
            /// <param name="handle">Handle du bouton.</param>
            /// <returns>Indique si le message a été posté.</returns>
            public bool Click(IntPtr handle) { return PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); }
            /// <summary>Retourne l’état de sélection d’un bouton à cocher.</summary>
            /// <param name="handle">Handle du contrôle.</param>
            /// <returns><see langword="true"/> si le bouton est coché.</returns>
            public bool Checked(IntPtr handle) { return SendMessageInt(handle, 0x00F0, IntPtr.Zero, IntPtr.Zero).ToInt32() == 1; }
            /// <summary>Sélectionne tout le texte du champ et saisit la nouvelle valeur.</summary>
            /// <param name="handle">Handle du champ texte.</param>
            /// <param name="value">Texte à saisir.</param>
            public void Replace(IntPtr handle, string value)
            {
                SendMessageInt(handle, 0x00B1, IntPtr.Zero, new IntPtr(-1));
                SendMessageText(handle, 0x00C2, new IntPtr(1), value);
            }
            /// <summary>Attend la durée en millisecondes demandée.</summary>
            /// <param name="milliseconds">Durée de pause.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
            /// <summary>Lit le message accessible du dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Texte accessible.</returns>
            public string Message(IntPtr dialog) { return AccessibleDialogMessage(dialog); }
            /// <summary>Ferme le dialogue après une interaction.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            public void Close(IntPtr dialog) { CloseDialog(dialog); }
            /// <summary>Recherche la fenêtre racine du VBE.</summary>
            /// <returns>Handle de la fenêtre, ou zéro si elle n’est pas trouvée.</returns>
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            /// <summary>Retourne les handles enfants d’une fenêtre racine.</summary>
            /// <param name="root">Handle racine.</param>
            /// <returns>Handles des fenêtres enfants visibles.</returns>
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            /// <summary>Recherche un volet par l’un des noms possibles.</summary>
            /// <param name="panes">Handles de volets candidats.</param>
            /// <param name="names">Noms possibles du volet.</param>
            /// <returns>Handle du volet, ou zéro.</returns>
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            /// <summary>Lit les lignes accessibles du volet.</summary>
            /// <param name="handle">Handle du volet.</param>
            /// <returns>Instantané du contenu.</returns>
            public object List(IntPtr handle) { return ReadList(handle); }
            /// <summary>Recherche les lignes de surveillance correspondant à l’expression et au contexte.</summary>
            /// <param name="pane">Handle du volet Watches.</param>
            /// <param name="expression">Expression exacte de surveillance.</param>
            /// <param name="context">Contexte facultatif d’expression.</param>
            /// <returns>Nombre de lignes correspondantes et conserve les éléments accessibles correspondants.</returns>
            public int WatchMatches(IntPtr pane, string expression, string context)
            {
                lastRows = MatchingWatchRows(pane, expression, context);
                return lastRows.Count;
            }
            /// <summary>Sélectionne la ligne de surveillance unique trouvée par la recherche précédente.</summary>
            /// <param name="pane">Volet Watches contenant la ligne.</param>
            /// <param name="expression">Expression demandée.</param>
            /// <param name="context">Contexte facultatif.</param>
            /// <returns><see langword="true"/> si le motif de sélection UI Automation est disponible et appliqué.</returns>
            public bool SelectWatchRow(IntPtr pane, string expression, string context)
            {
                AutomationElement row = lastRows.Single();
                if (!row.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern)) return false;
                ((SelectionItemPattern)pattern).Select();
                row.SetFocus();
                return true;
            }
        }

        /// <summary>Enfant MSAA d’une boîte de signature avec index, rôle et nom accessibles.</summary>
        internal sealed class SignatureChild
        {
            /// <summary>Index de l’enfant dans l’arbre accessible MSAA.</summary>
            public int Index;
            /// <summary>Rôle MSAA de l’enfant.</summary>
            public int Role;
            /// <summary>Nom accessible de l’enfant.</summary>
            public string Name;
        }

        /// <summary>Abstraction injectable de lecture et de fermeture d’une boîte de signature VBE.</summary>
        internal interface ISignatureProbe
        {
            /// <summary>Recherche la boîte de signature.</summary>
            /// <returns>Handle du dialogue ou zéro.</returns>
            IntPtr Dialog();
            /// <summary>Énumère les enfants accessibles du dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Enfants MSAA lisibles.</returns>
            IList<SignatureChild> Children(IntPtr dialog);
            /// <summary>Active l’action MSAA d’annulation de l’enfant indiqué.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <param name="index">Index de l’enfant.</param>
            void Cancel(IntPtr dialog, int index);
            /// <summary>Ferme le dialogue de signature.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            void Close(IntPtr dialog);
            /// <summary>Attend le délai indiqué.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            void Pause(int milliseconds);
        }

        /// <summary>Lit et ferme le dialogue de signature au moyen de l’interface d’accessibilité native.</summary>
        private sealed class NativeSignatureProbe : ISignatureProbe
        {
            /// <summary>Interface MSAA du dialogue ouvert.</summary>
            private Accessibility.IAccessible accessible;
            /// <summary>Recherche le dialogue natif de signature.</summary>
            /// <returns>Handle du dialogue ou zéro.</returns>
            public IntPtr Dialog() { return FindSignatureDialog(); }
            /// <summary>Lit les enfants MSAA disponibles dans le dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Index, rôle et nom des enfants accessibles.</returns>
            public IList<SignatureChild> Children(IntPtr dialog)
            {
                accessible = SignatureAccessible(dialog);
                var children = new List<SignatureChild>();
                for (int index = 1; index <= Math.Min(accessible.accChildCount, 64); index++)
                {
                    try
                    {
                        children.Add(new SignatureChild { Index = index, Name = accessible.get_accName(index),
                            Role = Convert.ToInt32(accessible.get_accRole(index)) });
                    }
                    catch { /* A single inaccessible MSAA child is skipped. */ }
                }
                return children;
            }
            /// <summary>Déclenche l’action accessible de l’enfant, utilisée pour annuler sans clic coordonné.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <param name="index">Index MSAA de l’action.</param>
            public void Cancel(IntPtr dialog, int index) { accessible.accDoDefaultAction(index); }
            /// <summary>Ferme le dialogue de signature par message natif.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            public void Close(IntPtr dialog) { PostMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero); }
            /// <summary>Attend le délai indiqué.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        /// <summary>Valeur observée d’un contrôle dans l’onglet des options VBE.</summary>
        internal sealed class OptionsControl
        {
            /// <summary>Nom accessible du contrôle.</summary>
            public string Name;
            /// <summary>Type de contrôle UI Automation.</summary>
            public string Type;
            /// <summary>Valeur courante lue, si elle est disponible.</summary>
            public object Value;
            /// <summary>Erreur de lecture du contrôle, si présente.</summary>
            public string Error;
            /// <summary>Indique si le contrôle est visible.</summary>
            public bool Visible = true;
            /// <summary>Indique si le contrôle est activé.</summary>
            public bool Enabled = true;
        }

        /// <summary>Choix radio ou case lu dans les options de débogage.</summary>
        internal sealed class OptionsChoice
        {
            /// <summary>Nom accessible du choix.</summary>
            public string Name;
            /// <summary>Indique si ce choix est sélectionné.</summary>
            public bool Selected;
            /// <summary>Indique si l’état du choix a pu être lu.</summary>
            public bool Readable = true;
        }

        /// <summary>Abstraction injectable de lecture du dialogue natif Tools &gt; Options.</summary>
        internal interface IOptionsProbe
        {
            /// <summary>Ouvre ou localise le dialogue Options.</summary>
            /// <returns>Handle du dialogue, ou zéro.</returns>
            IntPtr Dialog();
            /// <summary>Énumère les titres d’onglet du dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Noms visibles des onglets.</returns>
            IList<string> Tabs(IntPtr dialog);
            /// <summary>Lit les contrôles de l’onglet sélectionné par index.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <param name="tabIndex">Index d’onglet.</param>
            /// <returns>Valeurs observées des contrôles.</returns>
            IList<OptionsControl> Controls(IntPtr dialog, int tabIndex);
            /// <summary>Lit les choix du réglage de capture des erreurs.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Choix et état de sélection observés.</returns>
            IList<OptionsChoice> ErrorChoices(IntPtr dialog);
            /// <summary>Ferme le dialogue sans appliquer de modification.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            void Close(IntPtr dialog);
            /// <summary>Attend un délai entre les lectures d’interface.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            void Pause(int milliseconds);
        }

        /// <summary>Implémente la lecture du dialogue Options avec UI Automation.</summary>
        private sealed class NativeOptionsProbe : IWritableOptionsProbe
        {
            /// <summary>Racine UI Automation du dialogue.</summary>
            private AutomationElement root;
            /// <summary>Onglets UI Automation du dialogue.</summary>
            private AutomationElementCollection tabItems;
            /// <summary>Recherche le dialogue VBE des options.</summary>
            /// <returns>Handle du dialogue ou zéro.</returns>
            public IntPtr Dialog() { return FindDialog("Options"); }
            /// <summary>Lit les noms des onglets accessibles du dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Noms des onglets dans leur ordre d’affichage.</returns>
            public IList<string> Tabs(IntPtr dialog)
            {
                root = AutomationElement.FromHandle(dialog);
                tabItems = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                return tabItems.Cast<AutomationElement>().Select(tab => tab.Current.Name).ToArray();
            }
            /// <summary>Lit les valeurs des contrôles UI Automation de l’onglet demandé.</summary>
            /// <param name="dialog">Handle du dialogue Options.</param>
            /// <param name="tabIndex">Index d’onglet sélectionné.</param>
            /// <returns>Contrôles et valeurs, y compris les erreurs individuelles de lecture.</returns>
            public IList<OptionsControl> Controls(IntPtr dialog, int tabIndex)
            {
                object tabPattern;
                if (!tabItems[tabIndex].TryGetCurrentPattern(SelectionItemPattern.Pattern, out tabPattern))
                    throw new InvalidOperationException("A native VBE Options tab is unreadable.");
                ((SelectionItemPattern)tabPattern).Select();
                PauseNative(75);
                var descendants = root.FindAll(TreeScope.Descendants,
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Slider),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)));
                if (descendants.Count > 2000)
                    throw new InvalidOperationException("The Options dialog has too many controls to inspect safely: " +
                        descendants.Count + ".");
                var controls = new List<OptionsControl>();
                for (int index = 0; index < descendants.Count; index++)
                {
                    AutomationElement element = descendants[index];
                    try
                    {
                        ControlType kind = element.Current.ControlType;
                        var control = new OptionsControl { Name = element.Current.Name,
                            Type = kind.ProgrammaticName, Visible = !element.Current.IsOffscreen,
                            Enabled = element.Current.IsEnabled };
                        if (!control.Visible || !control.Enabled) { controls.Add(control); continue; }
                        try
                        {
                            if (kind == ControlType.CheckBox &&
                                element.TryGetCurrentPattern(TogglePattern.Pattern, out object toggle))
                                control.Value = ((TogglePattern)toggle).Current.ToggleState.ToString();
                            else if ((kind == ControlType.RadioButton || kind == ControlType.ListItem) &&
                                element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
                                control.Value = ((SelectionItemPattern)selection).Current.IsSelected;
                            else if ((kind == ControlType.Edit || kind == ControlType.ComboBox) &&
                                !element.Current.IsPassword &&
                                element.TryGetCurrentPattern(ValuePattern.Pattern, out object input))
                                control.Value = ((ValuePattern)input).Current.Value;
                            else if (kind == ControlType.Slider &&
                                element.TryGetCurrentPattern(RangeValuePattern.Pattern, out object slider))
                                control.Value = ((RangeValuePattern)slider).Current.Value;
                        }
                        catch (Exception ex) { control.Error = ex.Message; }
                        controls.Add(control);
                    }
                    catch (ElementNotAvailableException) { }
                }
                return controls;
            }
            /// <summary>Écrit un contrôle unique de l’onglet par son interface native accessible.</summary>
            public void Write(IntPtr dialog, int tabIndex, string name, string type, object value)
            {
                Controls(dialog, tabIndex);
                var candidates = root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name))
                    .Cast<AutomationElement>().Where(x => !x.Current.IsOffscreen && x.Current.IsEnabled && x.Current.ControlType.ProgrammaticName == type).ToArray();
                if (candidates.Length != 1) throw new InvalidOperationException("The exact option control is absent or ambiguous.");
                var selected = candidates[0];
                if (type == "ControlType.CheckBox" && selected.TryGetCurrentPattern(TogglePattern.Pattern, out object toggle))
                {
                    var pattern = (TogglePattern)toggle;
                    if (pattern.Current.ToggleState == ToggleState.Indeterminate) throw new InvalidOperationException("An indeterminate option is not writable.");
                    bool desired = (bool)value;
                    if ((pattern.Current.ToggleState == ToggleState.On) != desired) pattern.Toggle();
                }
                else if (type == "ControlType.RadioButton" && selected.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
                    ((SelectionItemPattern)selection).Select();
                else if (type == "ControlType.Edit" && !selected.Current.IsPassword && selected.TryGetCurrentPattern(ValuePattern.Pattern, out object input))
                {
                    var pattern = (ValuePattern)input;
                    if (pattern.Current.IsReadOnly) throw new InvalidOperationException("The option is read-only.");
                    pattern.SetValue(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));
                }
                else throw new InvalidOperationException("The option has no supported writable pattern.");
                PauseNative(100);
            }
            /// <summary>Demande la validation par le bouton natif IDOK, sans raccourci clavier.</summary>
            public void Accept(IntPtr dialog)
            {
                IntPtr ok = GetDlgItem(dialog, 1);
                if (ok == IntPtr.Zero || ClassName(ok) != "Button" || !OptionsWindowEnabled(ok) || !PostMessage(ok, 0x00F5, IntPtr.Zero, IntPtr.Zero))
                    throw new InvalidOperationException("The native Options OK button is unavailable.");
            }

            /// <summary>Lit les choix disponibles pour le réglage de capture d’erreurs.</summary>
            /// <param name="dialog">Handle du dialogue Options.</param>
            /// <returns>Libellés, états de sélection et disponibilité de lecture.</returns>
            public IList<OptionsChoice> ErrorChoices(IntPtr dialog)
            {
                AutomationElement options = AutomationElement.FromHandle(dialog);
                var generalCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new OrCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Général"),
                        new PropertyCondition(AutomationElement.NameProperty, "General")));
                AutomationElementCollection tabs = options.FindAll(TreeScope.Descendants, generalCondition);
                if (tabs.Count != 1 || !tabs[0].TryGetCurrentPattern(SelectionItemPattern.Pattern, out object tabPattern))
                    throw new InvalidOperationException("The native General options tab is unavailable.");
                ((SelectionItemPattern)tabPattern).Select();
                var radios = options.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton));
                var result = new List<OptionsChoice>();
                for (int index = 0; index < radios.Count; index++)
                {
                    AutomationElement radio = radios[index];
                    var choice = new OptionsChoice { Name = radio.Current.Name };
                    choice.Readable = radio.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern);
                    if (choice.Readable) choice.Selected = ((SelectionItemPattern)pattern).Current.IsSelected;
                    result.Add(choice);
                }
                return result;
            }
            /// <summary>Ferme le dialogue sans appliquer de choix.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            public void Close(IntPtr dialog) { CloseDialog(dialog); }
            /// <summary>Attend avant la lecture UI Automation suivante.</summary>
            /// <param name="milliseconds">Durée de l’attente.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        /// <summary>Abstraction injectable de lecture et d’envoi de texte à la fenêtre Immediate.</summary>
        internal interface IImmediateProbe
        {
            /// <summary>Retourne la fenêtre racine VBE.</summary>
            /// <returns>Handle racine, ou zéro.</returns>
            IntPtr VbeRoot();
            /// <summary>Énumère les fenêtres enfants.</summary>
            /// <param name="root">Handle racine.</param>
            /// <returns>Handles enfants.</returns>
            List<IntPtr> Children(IntPtr root);
            /// <summary>Recherche un volet par ses noms possibles.</summary>
            /// <param name="panes">Fenêtres candidates.</param>
            /// <param name="names">Noms de volet.</param>
            /// <returns>Handle trouvé, ou zéro.</returns>
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            /// <summary>Prépare le volet Immediate pour l’exécution d’une commande.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns>Texte observé avant exécution.</returns>
            string Prepare(IntPtr pane);
            /// <summary>Lit le texte du volet Immediate.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns>Texte accessible observé.</returns>
            string Text(IntPtr pane);
            /// <summary>Poste le caractère demandé dans le volet.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <param name="character">Caractère à saisir.</param>
            /// <returns><see langword="true"/> si le message est posté.</returns>
            bool PostChar(IntPtr pane, char character);
            /// <summary>Poste la séquence de touche Entrée.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns><see langword="true"/> si les messages sont postés.</returns>
            bool PostEnter(IntPtr pane);
            /// <summary>Attend l’exécution dans l’interface du VBE.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            void Pause(int milliseconds);
        }

        /// <summary>Implémente l’envoi de commandes texte dans le volet Immediate via les messages clavier natifs.</summary>
        private sealed class NativeImmediateProbe : IImmediateProbe
        {
            /// <summary>Recherche la fenêtre racine du VBE.</summary>
            /// <returns>Handle racine, ou zéro si absent.</returns>
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            /// <summary>Énumère les fenêtres enfants d’une racine.</summary>
            /// <param name="root">Handle parent.</param>
            /// <returns>Handles enfants.</returns>
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            /// <summary>Recherche le volet Immediate parmi les fenêtres candidates.</summary>
            /// <param name="panes">Handles candidats.</param>
            /// <param name="names">Noms possibles du volet.</param>
            /// <returns>Handle du volet ou zéro.</returns>
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            /// <summary>Place le focus dans le volet Immediate et lit son texte initial.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns>Texte observable après préparation.</returns>
            public string Prepare(IntPtr pane)
            {
                AutomationElement document = ImmediateDocument(pane);
                var pattern = (TextPattern)document.GetCurrentPattern(TextPattern.Pattern);
                string before = pattern.DocumentRange.GetText(-1);
                document.SetFocus();
                TextPatternRange atEnd = pattern.DocumentRange.Clone();
                atEnd.MoveEndpointByRange(TextPatternRangeEndpoint.Start, pattern.DocumentRange,
                    TextPatternRangeEndpoint.End);
                atEnd.Select();
                return before;
            }
            /// <summary>Lit le texte accessible courant du volet Immediate.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns>Texte affiché.</returns>
            public string Text(IntPtr pane) { return ImmediateText(pane); }
            /// <summary>Poste un message de caractère dans le volet.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <param name="character">Caractère à transmettre.</param>
            /// <returns><see langword="true"/> si le message a été posté.</returns>
            public bool PostChar(IntPtr pane, char character)
            { return PostMessage(pane, WmChar, new IntPtr(character), IntPtr.Zero); }
            /// <summary>Poste l’appui puis le relâchement de la touche Entrée.</summary>
            /// <param name="pane">Handle du volet.</param>
            /// <returns><see langword="true"/> si les messages ont été postés.</returns>
            public bool PostEnter(IntPtr pane)
            {
                return PostMessage(pane, WmKeyDown, new IntPtr(VkReturn), IntPtr.Zero) &&
                    PostMessage(pane, WmKeyUp, new IntPtr(VkReturn), IntPtr.Zero);
            }
            /// <summary>Attend la durée indiquée pour laisser le VBE traiter la saisie.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        /// <summary>Implémente la lecture des volets et dialogues par Win32, UI Automation et MSAA.</summary>
        private sealed class NativeProbe : INativeProbe
        {
            /// <summary>Recherche la fenêtre racine du VBE.</summary>
            /// <returns>Handle de la fenêtre, ou zéro.</returns>
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            /// <summary>Retourne les fenêtres enfants d’une racine.</summary>
            /// <param name="root">Handle parent.</param>
            /// <returns>Handles enfants.</returns>
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            /// <summary>Recherche un volet parmi les fenêtres candidates.</summary>
            /// <param name="panes">Fenêtres à examiner.</param>
            /// <param name="names">Titres possibles.</param>
            /// <returns>Handle du volet ou zéro.</returns>
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            /// <summary>Lit les lignes accessibles d’un volet.</summary>
            /// <param name="handle">Handle du volet.</param>
            /// <returns>Instantané de la liste.</returns>
            public object List(IntPtr handle) { return ReadList(handle); }
            /// <summary>Lit l’éditeur Immediate.</summary>
            /// <param name="handle">Handle du volet Immediate.</param>
            /// <returns>Instantané de texte.</returns>
            public object Immediate(IntPtr handle) { return ReadImmediate(handle); }
            /// <summary>Lit la pile d’appels associée au volet Locals.</summary>
            /// <param name="locals">Handle du volet Locals.</param>
            /// <returns>État de la pile d’appels.</returns>
            public object CallStack(IntPtr locals) { return ReadCallStack(locals); }
            /// <summary>Recherche un dialogue par ses titres possibles.</summary>
            /// <param name="titles">Titres à rechercher.</param>
            /// <returns>Handle du dialogue ou zéro.</returns>
            public IntPtr Dialog(params string[] titles) { return FindDialog(titles); }
            /// <summary>Énumère les contrôles enfants d’un dialogue natif.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Contrôles avec handle, classe, texte et visibilité.</returns>
            public List<NativeControl> DialogControls(IntPtr dialog)
            {
                var controls = new List<NativeControl>();
                EnumChildWindows(dialog, (handle, parameter) => {
                    controls.Add(new NativeControl { Handle = handle, Kind = ClassName(handle),
                        Text = WindowText(handle), Visible = IsWindowVisible(handle) });
                    return true;
                }, IntPtr.Zero);
                return controls;
            }
            /// <summary>Lit le message accessible d’un dialogue.</summary>
            /// <param name="dialog">Handle du dialogue.</param>
            /// <returns>Message accessible.</returns>
            public string DialogMessage(IntPtr dialog) { return AccessibleDialogMessage(dialog); }
            /// <summary>Poste une activation standard du contrôle.</summary>
            /// <param name="handle">Handle à activer.</param>
            /// <returns><see langword="true"/> si le message est posté.</returns>
            public bool Click(IntPtr handle) { return PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); }
            /// <summary>Vérifie la visibilité actuelle d’une fenêtre.</summary>
            /// <param name="handle">Handle à tester.</param>
            /// <returns><see langword="true"/> si visible.</returns>
            public bool Visible(IntPtr handle) { return IsWindowVisible(handle); }
            /// <summary>Attend la durée indiquée.</summary>
            /// <param name="milliseconds">Durée en millisecondes.</param>
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        /// <summary>Capture les volets visibles Locals, Watches et Immediate, avec la pile d’appels optionnelle.</summary>
        /// <param name="includeCallStack">Ouvre temporairement la pile d’appels pour la lire lorsque la valeur est vraie.</param>
        /// <returns>Instantané des volets accessibles et limites de la capture.</returns>
        /// <exception cref="InvalidOperationException">La fenêtre racine VBE n’est pas présente dans le processus courant.</exception>
        public static object Capture(bool includeCallStack)
        {
            return Capture(includeCallStack, new NativeProbe());
        }

        /// <summary>Capture les volets avec une sonde injectable pour séparer l’orchestration du transport natif.</summary>
        /// <param name="includeCallStack">Inclut la pile d’appels si vrai.</param>
        /// <param name="native">Sonde des fenêtres natives.</param>
        /// <returns>Instantané des volets accessibles.</returns>
        /// <exception cref="ArgumentNullException">La sonde est nulle.</exception>
        /// <exception cref="InvalidOperationException">La fenêtre racine du VBE est absente.</exception>
        internal static object Capture(bool includeCallStack, INativeProbe native)
        {
            if (native == null) throw new ArgumentNullException(nameof(native));
            IntPtr root = native.VbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open in this host process.");
            var panes = native.Children(root);
            IntPtr locals = native.Pane(panes, "Variables locales", "Locals");
            IntPtr watches = native.Pane(panes, "Espions", "Watch", "Watches");
            IntPtr immediate = native.Pane(panes, "Exécution", "Immediate");
            object stack = includeCallStack ? native.CallStack(locals) : null;
            return new {
                HostProcessId = Process.GetCurrentProcess().Id,
                Locals = native.List(locals),
                Watches = native.List(watches),
                Immediate = native.Immediate(immediate),
                CallStack = stack,
                Limits = "Native UI accessibility is observed only for visible panes. A missing pane or unavailable value is not an empty debugger collection. Breakpoints and exception state are not exposed by this snapshot."
            };
        }

        // The Compile command can open a modal native diagnostic. Its UI-thread
        // Execute call cannot be awaited with Control.Invoke in that case.
        /// <summary>Vérifie qu’aucun dialogue d’erreur de compilation n’est encore ouvert.</summary>
        /// <exception cref="InvalidOperationException">Un dialogue de compilation reste visible.</exception>
        public static void EnsureNoCompileDialog()
        {
            EnsureNoCompileDialog(new NativeProbe());
        }

        /// <summary>Vérifie avec une sonde fournie l’absence de dialogue de compilation.</summary>
        /// <param name="native">Sonde des dialogues natifs.</param>
        /// <exception cref="InvalidOperationException">Un dialogue de compilation reste visible.</exception>
        internal static void EnsureNoCompileDialog(INativeProbe native)
        {
            if (native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications", "Microsoft Visual Basic") != IntPtr.Zero)
                throw new InvalidOperationException("A native VBE dialog is already open; compilation was not started.");
        }

        /// <summary>Lit le message et les boutons d’un dialogue de diagnostic VBA visible sans le fermer.</summary>
        /// <returns>Message, commandes disponibles et état du dialogue observé.</returns>
        public static object ReadDebugDialog()
        {
            return ReadDebugDialog(new NativeProbe());
        }

        /// <summary>Lit un dialogue de diagnostic au moyen de la sonde spécifiée.</summary>
        /// <param name="native">Sonde native d’interface.</param>
        /// <returns>Texte accessible et boutons observés.</returns>
        internal static object ReadDebugDialog(INativeProbe native)
        {
            IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
            if (dialog == IntPtr.Zero)
                return new { Available = false, Diagnostic = (string)null,
                    Buttons = new string[0], Error = (string)null };
            var messages = new List<string>();
            var buttons = new List<string>();
            foreach (var control in native.DialogControls(dialog))
            {
                if (!control.Visible || string.IsNullOrWhiteSpace(control.Text)) continue;
                string title = control.Text;
                string kind = control.Kind;
                if (kind == "Static") messages.Add(title);
                else if (kind == "Button") buttons.Add(title);
            }
            if (messages.Count != 1)
                return new { Available = true, Diagnostic = (string)null,
                    Buttons = buttons.ToArray(), Error = "Expected one native diagnostic message; found " + messages.Count + "." };
            return new { Available = true, Diagnostic = messages[0],
                Buttons = buttons.ToArray(), Error = (string)null };
        }

        /// <summary>Échoue si le dialogue natif des options de débogage est ouvert.</summary>
        /// <exception cref="InvalidOperationException">Le dialogue d’options est visible.</exception>
        public static void EnsureNoDebugOptionsDialog()
        {
            EnsureNoDebugOptionsDialog(new NativeProbe());
        }

        /// <summary>Vérifie avec la sonde fournie l’absence du dialogue d’options de débogage.</summary>
        /// <param name="native">Sonde native d’interface.</param>
        /// <exception cref="InvalidOperationException">Le dialogue d’options est visible.</exception>
        internal static void EnsureNoDebugOptionsDialog(INativeProbe native)
        {
            if (native.Dialog("Options") != IntPtr.Zero)
                throw new InvalidOperationException("A VBE Options dialog is already open; the add-in will not close a user-owned dialog.");
        }

        /// <summary>Échoue si un dialogue de signature de projet est déjà visible.</summary>
        /// <exception cref="InvalidOperationException">Un dialogue de signature est ouvert.</exception>
        public static void EnsureNoSignatureDialog()
        {
            if (FindSignatureDialog() != IntPtr.Zero)
                throw new InvalidOperationException("A VBE Digital Signature dialog is already open; the add-in will not close a user-owned dialog.");
        }

        /// <summary>Lit le contenu accessible du dialogue de signature sans appliquer ni enregistrer de choix.</summary>
        /// <param name="project">Nom du projet affiché dans le résultat.</param>
        /// <returns>État du dialogue et choix de certificat observés.</returns>
        public static object ReadSignatureDialog(string project)
        {
            return ReadSignatureDialog(project, new NativeSignatureProbe());
        }

        /// <summary>Lit le dialogue de signature par l’abstraction MSAA injectée.</summary>
        /// <param name="project">Nom du projet concerné.</param>
        /// <param name="native">Sonde d’accessibilité du dialogue.</param>
        /// <returns>Contenu accessible observé.</returns>
        internal static object ReadSignatureDialog(string project, ISignatureProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Digital Signature dialog did not open.");
            var labels = new List<string>();
            var buttons = new List<string>();
            bool cancelled = false;
            try
            {
                int cancelIndex = 0;
                foreach (SignatureChild child in native.Children(dialog))
                {
                    string name = child.Name;
                    int role = child.Role;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (role == 43)
                    {
                        buttons.Add(name);
                        if (string.Equals(name, "Annuler", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(name, "Cancel", StringComparison.OrdinalIgnoreCase))
                            cancelIndex = child.Index;
                    }
                    else if (role == 41) labels.Add(name);
                }
                if (cancelIndex == 0)
                    throw new InvalidOperationException("The native Digital Signature dialog has no accessible Cancel button.");
                native.Cancel(dialog, cancelIndex);
                cancelled = true;
            }
            finally
            {
                if (!cancelled && native.Dialog() == dialog)
                    native.Close(dialog); // WM_CLOSE on our exact dialog
            }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The signature dialog was read but did not close.");
            string currentCertificate = CertificateBeforeHeading(labels,
                "Signature actuelle du projet VBA", "The VBA project is currently signed as");
            string signAsCertificate = CertificateBeforeHeading(labels, "Signer en tant que", "Sign as");
            return new { Project = project, CurrentCertificate = currentCertificate,
                SignAsCertificate = signAsCertificate, Labels = labels.ToArray(), Buttons = buttons.ToArray(),
                DialogClosed = true, Verification = "NativeSignatureDialogReadback",
                Limit = "Labels reflect the native dialog; no certificate was selected, assigned, removed or cryptographically validated." };
        }

        // Windows' protected certificate picker does not expose its buttons to
        // UIA/MSAA. The user confirms the named certificate there; the add-in
        // checks VBE's readback and completes only its own native VBE dialog.
        /// <summary>Sélectionne le certificat demandé dans le dialogue natif et vérifie la signature rapportée par le VBE.</summary>
        /// <param name="project">Nom du projet à signer.</param>
        /// <param name="thumbprint">Empreinte du certificat exact à sélectionner.</param>
        /// <param name="certificateName">Nom du certificat affiché attendu.</param>
        /// <param name="unsignedVerified">Indique si l’absence de signature initiale a déjà été vérifiée.</param>
        /// <returns>État de signature observé après fermeture du dialogue.</returns>
        /// <exception cref="InvalidOperationException">Le certificat ou la boîte native ne peut pas être confirmé.</exception>
        public static object CompleteProjectSignature(string project, string thumbprint, string certificateName,
            bool unsignedVerified)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { PauseNative(50); dialog = FindSignatureDialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Digital Signature dialog did not open.");
            bool completed = false;
            try
            {
                Accessibility.IAccessible root = SignatureAccessible(dialog);
                var initial = SignatureLabels(root);
                string current = CertificateBeforeHeading(initial,
                    "Signature actuelle du projet VBA", "The VBA project is currently signed as");
                string signAs = CertificateBeforeHeading(initial, "Signer en tant que", "Sign as");
                if ((!unsignedVerified && !IsNoCertificate(current)) ||
                    (!IsNoCertificate(signAs) && !string.Equals(signAs, certificateName, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("This command adds the first signature only; the project already has a certificate selected.");
                bool chosen = string.Equals(signAs, certificateName, StringComparison.OrdinalIgnoreCase) ||
                    (unsignedVerified && IsNoCertificate(signAs) &&
                     string.Equals(current, certificateName, StringComparison.OrdinalIgnoreCase));
                bool pickerOpened = !chosen;
                if (!chosen) InvokeSignatureButton(root, "Choisir...", "Choose...");
                for (int attempt = 0; !chosen && attempt < 2400; attempt++)
                {
                    PauseNative(50);
                    if (FindSignatureDialog() != dialog)
                        throw new InvalidOperationException("The native VBE signature dialog closed during certificate selection.");
                    if (attempt % 4 != 0) continue;
                    string selected = CertificateBeforeHeading(SignatureLabels(root), "Signer en tant que", "Sign as");
                    if (string.Equals(selected, certificateName, StringComparison.OrdinalIgnoreCase))
                    { chosen = true; break; }
                    if (!IsNoCertificate(selected))
                        throw new InvalidOperationException("A different certificate was selected; the signature was cancelled.");
                }
                if (!chosen)
                    throw new TimeoutException("The certificate was not confirmed in Windows Security within two minutes.");
                InvokeSignatureButton(root, "OK");
                for (int attempt = 0; attempt < 60; attempt++)
                {
                    if (FindSignatureDialog() == IntPtr.Zero) { completed = true; break; }
                    PauseNative(50);
                }
                if (!completed) throw new InvalidOperationException("The native VBE signature dialog did not close after OK.");
                return new { Project = project, CertificateThumbprint = thumbprint,
                    CertificateName = certificateName, SignatureAssigned = true,
                    Verification = "NativeVbeCertificateReadbackAndDialogClose",
                    SelectionSource = pickerOpened ? "WindowsCertificatePicker" : "NativeVbeExistingCertificate",
                    CertificatePicker = pickerOpened
                        ? "The Windows certificate picker returned the named certificate; the add-in did not interact with its protected controls."
                        : "The named certificate was already associated with the project in the native VBE dialog." };
            }
            finally
            {
                if (!completed && FindSignatureDialog() == dialog)
                    PostMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero);
            }
        }

        /// <summary>Obtient l’interface MSAA du dialogue de signature et vérifie la disponibilité des enfants.</summary>
        /// <param name="dialog">Handle du dialogue.</param>
        /// <returns>Interface accessible du dialogue.</returns>
        private static Accessibility.IAccessible SignatureAccessible(IntPtr dialog)
        {
            object accessible;
            Guid iid = IidAccessible;
            int hr = AccessibleObjectFromWindow(dialog, ObjidClient, ref iid, out accessible);
            if (hr != 0 || !(accessible is Accessibility.IAccessible))
                throw new COMException("The native signature dialog is not accessible through MSAA.", hr);
            return (Accessibility.IAccessible)accessible;
        }

        /// <summary>Collecte les libellés accessibles des enfants d’un dialogue de signature.</summary>
        /// <param name="root">Racine MSAA du dialogue.</param>
        /// <returns>Libellés non vides dans l’ordre d’énumération.</returns>
        private static List<string> SignatureLabels(Accessibility.IAccessible root)
        {
            var labels = new List<string>();
            for (int index = 1; index <= Math.Min(root.accChildCount, 64); index++)
            {
                try
                {
                    if (Convert.ToInt32(root.get_accRole(index)) == 41)
                    {
                        string name = root.get_accName(index);
                        if (!string.IsNullOrWhiteSpace(name)) labels.Add(name);
                    }
                }
                catch (COMException) { }
            }
            return labels;
        }

        /// <summary>Reconnaît les libellés indiquant l’absence de certificat utilisable.</summary>
        /// <param name="name">Libellé accessible à vérifier.</param>
        /// <returns><see langword="true"/> si le choix décrit l’absence de certificat.</returns>
        private static bool IsNoCertificate(string name)
        {
            return name != null && (name.Equals("[Aucun certificat]", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("[No certificate]", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("[None]", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Active un bouton accessible du dialogue en exigeant une correspondance unique par libellé.</summary>
        /// <param name="root">Racine MSAA.</param>
        /// <param name="names">Libellés possibles du bouton.</param>
        /// <exception cref="InvalidOperationException">Le bouton n’est pas trouvé de façon unique.</exception>
        private static void InvokeSignatureButton(Accessibility.IAccessible root, params string[] names)
        {
            for (int index = 1; index <= Math.Min(root.accChildCount, 64); index++)
            {
                try
                {
                    if (Convert.ToInt32(root.get_accRole(index)) == 43 &&
                        names.Any(name => string.Equals(root.get_accName(index), name, StringComparison.OrdinalIgnoreCase)))
                    { root.accDoDefaultAction(index); return; }
                }
                catch (COMException) { }
            }
            throw new InvalidOperationException("The expected accessible signature button is unavailable: " + names[0]);
        }

        /// <summary>Extrait le nom de certificat qui précède un des titres donnés dans le dialogue.</summary>
        /// <param name="labels">Libellés accessibles du dialogue.</param>
        /// <param name="headings">Titres de section possibles.</param>
        /// <returns>Nom du certificat détecté, ou chaîne vide.</returns>
        private static string CertificateBeforeHeading(IList<string> labels, params string[] headings)
        {
            for (int index = 2; index < labels.Count; index++)
                if (headings.Any(heading => string.Equals(labels[index], heading, StringComparison.OrdinalIgnoreCase)) &&
                    (labels[index - 1].StartsWith("Nom du certificat", StringComparison.OrdinalIgnoreCase) ||
                     labels[index - 1].StartsWith("Certificate name", StringComparison.OrdinalIgnoreCase)))
                    return labels[index - 2];
            return null;
        }

        /// <summary>Recherche la fenêtre modale de signature associée à Excel ou au VBE.</summary>
        /// <returns>Handle du dialogue trouvé, ou zéro.</returns>
        private static IntPtr FindSignatureDialog()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid != currentPid || !IsWindowVisible(handle)) return true;
                string title = WindowText(handle);
                if (!string.Equals(title, "Signature numérique", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(title, "Digital Signature", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(title, "Signature électronique", StringComparison.OrdinalIgnoreCase)) return true;
                string kind = ClassName(handle);
                if (kind != "#32770" && !kind.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase)) return true;
                result = handle;
                return false;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Lit le réglage natif de capture des erreurs, puis ferme le dialogue sans changer les préférences.</summary>
        /// <returns>Choix sélectionné et liste des choix proposés par le VBE.</returns>
        public static object ReadDebugOptions()
        {
            return ReadDebugOptions(new NativeOptionsProbe());
        }

        /// <summary>Lit le réglage de capture d’erreurs avec la sonde d’options fournie.</summary>
        /// <param name="native">Sonde du dialogue natif.</param>
        /// <returns>Choix sélectionné et état de lecture.</returns>
        internal static object ReadDebugOptions(IOptionsProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The native VBE Options dialog did not open.");
            string selected = null;
            string[] choices = null;
            try
            {
                IList<OptionsChoice> radios = native.ErrorChoices(dialog);
                if (radios.Count != 3)
                    throw new InvalidOperationException("Expected three native error trapping choices; found " + radios.Count + ".");
                var names = new List<string>();
                foreach (OptionsChoice radio in radios)
                {
                    string name = radio.Name;
                    if (string.IsNullOrWhiteSpace(name) ||
                        !(name.StartsWith("Arrêt ", StringComparison.OrdinalIgnoreCase) ||
                          name.StartsWith("Break ", StringComparison.OrdinalIgnoreCase)) ||
                        !radio.Readable)
                        throw new InvalidOperationException("A native error trapping radio is unreadable.");
                    names.Add(name);
                    if (radio.Selected)
                    {
                        if (selected != null) throw new InvalidOperationException("Multiple error trapping choices appear selected.");
                        selected = name;
                    }
                }
                if (selected == null) throw new InvalidOperationException("No error trapping choice appears selected.");
                choices = names.ToArray();
            }
            finally { native.Close(dialog); }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The add-in read VBE Options but could not close its dialog.");
            return new { Scope = "VBE", ErrorTrapping = selected, Choices = choices,
                Verification = "NativeOptionsReadback", DialogClosed = true,
                Limit = "This is the currently displayed VBE-wide preference, not a diagnosis of an active runtime error." };
        }
        /// <summary>Lit les onglets et contrôles du dialogue Options sans enregistrer ni modifier les préférences.</summary>
        /// <returns>Valeurs visibles et erreurs individuelles de lecture.</returns>
        public static object ReadVbeOptions()
        {
            return ReadVbeOptions(new NativeOptionsProbe());
        }

        /// <summary>Lit le dialogue Options avec une sonde injectable et le ferme sans appliquer de modifications.</summary>
        /// <param name="native">Sonde du dialogue d’options.</param>
        /// <returns>Onglets et contrôles observés.</returns>
        internal static object ReadVbeOptions(IOptionsProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Options dialog did not open.");
            List<object> tabs;
            try { tabs = CaptureOptionsTabs(native, dialog); }
            finally { native.Close(dialog); }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The add-in read VBE Options but could not close its dialog.");
            return new { Scope = "VBE", Tabs = tabs, Count = tabs.Count,
                OptionsVersion = OptionsRevision(tabs), DialogClosed = true, Verification = "NativeOptionsReadback",
                Limit = "Only visible native controls were observed; no settings were changed." };
        }
        /// <summary>Exécute une ligne texte dans le volet Immediate visible en vérifiant le contenu avant et après.</summary>
        /// <param name="command">Ligne à saisir dans le VBE.</param>
        /// <returns>Texte observé avant et après l’exécution.</returns>
        public static object ExecuteImmediate(string command)
        {
            return ExecuteImmediate(command, new NativeImmediateProbe());
        }

        /// <summary>Exécute une ligne Immediate par une sonde injectable et retourne les états relus.</summary>
        /// <param name="command">Ligne à exécuter.</param>
        /// <param name="native">Sonde du volet Immediate.</param>
        /// <returns>Texte et résultat observés avant et après.</returns>
        internal static object ExecuteImmediate(string command, IImmediateProbe native)
        {
            if (string.IsNullOrWhiteSpace(command) || command.Length > 2048 ||
                command.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0 ||
                command.Any(character => char.IsControl(character)))
                throw new ArgumentException("Immediate command must be one nonempty printable line of at most 2048 characters.");
            IntPtr root = native.VbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open.");
            IntPtr pane = native.Pane(native.Children(root), "Exécution", "Immediate");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The Immediate window must be visible.");
            string before = native.Prepare(pane);
            foreach (char character in command)
                if (!native.PostChar(pane, character))
                    throw new InvalidOperationException("The native Immediate pane rejected a character message.");
            string echoed = null;
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(25);
                echoed = native.Text(pane);
                if (echoed == before + command || echoed == before + command + "\r\n") break;
            }
            if (echoed != before + command && echoed != before + command + "\r\n")
                throw new InvalidOperationException("The Immediate pane did not echo the exact command; Enter was not sent.");
            if (!native.PostEnter(pane))
                throw new InvalidOperationException("The native Immediate pane rejected Enter.");
            string after = echoed;
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(25);
                after = native.Text(pane);
                if (after != echoed) break;
            }
            return new { Command = command, TextBefore = before, TextAfter = after,
                OutputDelta = after.StartsWith(before, StringComparison.Ordinal)
                    ? after.Substring(before.Length) : (string)null,
                CommandEchoObserved = true,
                Verification = after != echoed ? "ImmediateTextChangedAfterEnter" : "Pending",
                VerificationPending = after == echoed,
                Limit = "Text changed after Enter, but this does not prove an arbitrary VBA statement had the intended side effect. Read debug_state and relevant values separately." };
        }

        /// <summary>Déplie ou replie une ligne des volets Locals ou Watches par son chemin hiérarchique.</summary>
        /// <param name="request">Requête contenant volet, chemin d’éléments et action.</param>
        /// <returns>Nombre observé d’enfants après l’action.</returns>
        /// <exception cref="InvalidOperationException">Le volet, la ligne ou l’action ne correspond pas de manière unique.</exception>
        public static object ChangeDebugItem(Request request)
        {
            if (request == null || (request.Pane != "locals" && request.Pane != "watches") ||
                (request.Action != "expand" && request.Action != "collapse") ||
                request.PathSegments == null || request.PathSegments.Length == 0 ||
                request.PathSegments.Length > 16 || request.PathSegments.Any(string.IsNullOrWhiteSpace))
                throw new ArgumentException("Pane, Action and nonempty PathSegments are required.");
            IntPtr root = FindVbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open.");
            IntPtr pane = request.Pane == "locals" ?
                FindPane(ChildWindows(root), "Variables locales", "Locals") :
                FindPane(ChildWindows(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The requested debug pane is not visible.");
            AutomationElement scope = AutomationElement.FromHandle(pane);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
            AutomationElementCollection rows = scope.FindAll(TreeScope.Descendants, condition);
            AutomationElement target = null;
            int matches = 0;
            for (int index = 0; index < rows.Count; index++)
            {
                if (!ItemPath(rows[index]).SequenceEqual(request.PathSegments, StringComparer.Ordinal)) continue;
                if (request.Pane == "watches" && !string.IsNullOrWhiteSpace(request.Context) &&
                    !string.Equals(WatchRootContext(rows[index]), request.Context, StringComparison.OrdinalIgnoreCase))
                    continue;
                target = rows[index]; matches++;
            }
            if (matches != 1)
                throw new InvalidOperationException("Expected one matching debug item; found " + matches + ".");
            object rawPattern;
            if (!target.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out rawPattern))
                throw new InvalidOperationException("The selected debug item cannot be expanded or collapsed.");
            var pattern = (ExpandCollapsePattern)rawPattern;
            if (request.Action == "expand") pattern.Expand();
            else pattern.Collapse();
            PauseNative(50);
            int childCount = target.FindAll(TreeScope.Children, condition).Count;
            bool verified = request.Action == "expand" ? childCount > 0 : childCount == 0;
            return new { request.Pane, request.Action, request.PathSegments,
                ChildCount = childCount, Verification = verified ? "Observed" : "Pending",
                VerificationPending = !verified,
                CountLimit = "ChildCount counts UIA-exposed children at this moment; long native trees may expose only a subset.",
                NextRead = "Call debug_windows to read the current hierarchical rows." };
        }

        /// <summary>Active le bouton demandé dans un dialogue de diagnostic VBA visible.</summary>
        /// <param name="request">Requête contenant les textes exacts du diagnostic et du bouton.</param>
        /// <returns>État du dialogue après l’action.</returns>
        public static object RespondDebugDialog(Request request)
        {
            return RespondDebugDialog(request, new NativeProbe());
        }

        /// <summary>Répond à un dialogue avec une sonde native injectable après vérification du texte affiché.</summary>
        /// <param name="request">Diagnostic et bouton exacts à activer.</param>
        /// <param name="native">Sonde des dialogues natifs.</param>
        /// <returns>Message et état du dialogue après activation.</returns>
        internal static object RespondDebugDialog(Request request, INativeProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Diagnostic) ||
                string.IsNullOrWhiteSpace(request.Button))
                throw new ArgumentException("Diagnostic and Button are required.");
            IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("No native VBE dialog is visible.");
            IntPtr target = IntPtr.Zero;
            string message = null;
            int matches = 0;
            foreach (var control in native.DialogControls(dialog))
            {
                if (!control.Visible) continue;
                string kind = control.Kind;
                string title = control.Text;
                if (kind == "Static" && !string.IsNullOrWhiteSpace(title)) message = title;
                if (kind == "Button" && string.Equals(title, request.Button, StringComparison.Ordinal))
                { target = control.Handle; matches++; }
            }
            if (!string.Equals(message, request.Diagnostic, StringComparison.Ordinal))
                throw new InvalidOperationException("The native diagnostic changed before the button could be activated.");
            if (!IsRecognizedDiagnostic(message))
                throw new InvalidOperationException("The visible VBE dialog is not a recognized VBA diagnostic.");
            if (matches != 1 || target == IntPtr.Zero)
                throw new InvalidOperationException("Expected exactly one matching native dialog button.");
            if (!native.Click(target))
                throw new InvalidOperationException("The native dialog button could not be activated.");
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(50);
                if (!native.Visible(dialog))
                    return new { Activated = true, request.Button, Diagnostic = message,
                        Verification = "DialogClosed", VerificationPending = false };
            }
            return new { Activated = true, request.Button, Diagnostic = message,
                Verification = "Pending", VerificationPending = true };
        }

        /// <summary>Attend une compilation et retourne le texte d’un dialogue natif de compilation s’il apparaît.</summary>
        /// <param name="completed">Signal indiquant que la commande de compilation s’est terminée.</param>
        /// <returns>Diagnostic observé, ou chaîne vide si aucun dialogue reconnu n’apparaît.</returns>
        public static string AwaitCompileDialog(ManualResetEventSlim completed)
        {
            return AwaitCompileDialog(completed, new NativeProbe());
        }

        /// <summary>Attend la fin de compilation avec une sonde injectable et lit le premier dialogue de diagnostic reconnu.</summary>
        /// <param name="completed">Signal de fin de compilation.</param>
        /// <param name="native">Sonde du dialogue natif.</param>
        /// <returns>Texte du diagnostic reconnu, ou chaîne vide.</returns>
        internal static string AwaitCompileDialog(ManualResetEventSlim completed, INativeProbe native)
        {
            for (int attempt = 0; attempt < 200; attempt++)
            {
                IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                    "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
                if (dialog != IntPtr.Zero)
                {
                    string diagnostic = native.DialogMessage(dialog);
                    IntPtr ok = native.DialogControls(dialog)
                        .Where(control => control.Kind == "Button" &&
                            string.Equals(control.Text, "OK", StringComparison.OrdinalIgnoreCase))
                        .Select(control => control.Handle).FirstOrDefault();
                    if (ok == IntPtr.Zero || !native.Click(ok))
                        throw new InvalidOperationException("The native compile diagnostic could not be dismissed.");
                    if (!completed.Wait(3000))
                        throw new TimeoutException("The Compile command did not return after its diagnostic closed.");
                    return diagnostic;
                }
                if (completed.IsSet)
                {
                    // Allow the VBE to surface a delayed diagnostic after Execute.
                    native.Pause(250);
                    dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                        "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
                    if (dialog == IntPtr.Zero) return null;
                }
                native.Pause(50);
            }
            throw new TimeoutException("The native Compile command did not finish within ten seconds.");
        }

        /// <summary>Indique si le message correspond à un diagnostic de compilation VBA reconnu.</summary>
        /// <param name="message">Texte de dialogue à classifier.</param>
        /// <returns><see langword="true"/> pour une erreur de compilation reconnue.</returns>
        internal static bool IsRecognizedDiagnostic(string message)
        {
            return string.Equals(message, "L'identificateur sous le curseur n'est pas reconnu", StringComparison.Ordinal) ||
                message != null && Regex.IsMatch(message, @"\AImpossible d'aller à '[^'\r\n]{1,255}' qui est caché\z") ||
                message != null && Regex.IsMatch(message,
                @"^(Erreur d'exécution|Run-time error|Erreur de compilation|Compile error)",
                RegexOptions.IgnoreCase);
        }

        /// <summary>Complète le dialogue Add Watch avec l’expression et le contexte demandés.</summary>
        /// <param name="request">Requête portant l’expression et l’emplacement du code.</param>
        /// <returns>Expression de surveillance ajoutée et état vérifié.</returns>
        public static object CompleteAddWatch(Request request)
        {
            return CompleteAddWatch(request, new NativeWatchProbe());
        }

        /// <summary>Remplit et valide Add Watch par l’abstraction native fournie.</summary>
        /// <param name="request">Expression et contexte attendus.</param>
        /// <param name="native">Sonde du dialogue et du volet Watches.</param>
        /// <returns>État du dialogue et résultat observé.</returns>
        internal static object CompleteAddWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Ajouter un espion", "Add Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Add Watch dialog did not open.");
            bool completed = false;
            try
            {
                IntPtr edit = native.Item(dialog, 4853);
                IntPtr project = native.Item(dialog, 4858);
                IntPtr module = native.Item(dialog, 4857);
                IntPtr procedure = native.Item(dialog, 4856);
                IntPtr ok = native.Item(dialog, 1);
                if (edit == IntPtr.Zero || project == IntPtr.Zero || module == IntPtr.Zero || ok == IntPtr.Zero)
                    throw new InvalidOperationException("The native Add Watch dialog controls changed.");
                string shownProject = native.Text(project);
                string shownModule = native.Text(module);
                string shownProcedure = procedure == IntPtr.Zero ? null : native.Text(procedure);
                if (!string.Equals(shownProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(shownModule, request.Module, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(request.Procedure) &&
                     !string.Equals(shownProcedure, request.Procedure, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Add Watch context changed: " + shownProject + "." + shownModule + "." + shownProcedure);
                int typeId = request.WatchType == "break_when_true" ? 4851 :
                    request.WatchType == "break_when_changed" ? 4852 : 4850;
                IntPtr typeButton = native.Item(dialog, typeId);
                if (typeButton == IntPtr.Zero || !native.Click(typeButton))
                    throw new InvalidOperationException("The native watch type option was unavailable.");
                native.Pause(30);
                if (!native.Checked(typeButton))
                    throw new InvalidOperationException("The native watch type option was not selected.");
                // EM_REPLACESEL triggers the VBE's edit notifications. WM_SETTEXT
                // alone changes the visible text but is rejected as an empty expression.
                native.Replace(edit, request.Expression);
                if (!string.Equals(native.Text(edit), request.Expression, StringComparison.Ordinal))
                    throw new InvalidOperationException("The watch expression was not reflected by the native edit control.");
                if (!native.Click(ok))
                    throw new InvalidOperationException("The Add Watch dialog refused its OK command.");
                IntPtr error = IntPtr.Zero;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    native.Pause(50);
                    dialog = native.Dialog("Ajouter un espion", "Add Watch");
                    if (dialog == IntPtr.Zero) { completed = true; break; }
                    error = native.Dialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
                    if (error != IntPtr.Zero) break;
                }
                if (error != IntPtr.Zero)
                {
                    string detail = native.Message(error);
                    native.Close(error);
                    throw new InvalidOperationException("VBE rejected the watch expression: " + detail);
                }
                if (!completed) throw new InvalidOperationException("The Add Watch dialog did not close after OK.");
                IntPtr root = native.VbeRoot();
                IntPtr watches = root == IntPtr.Zero ? IntPtr.Zero :
                    native.Pane(native.Children(root), "Espions", "Watch", "Watches");
                object watchState = native.List(watches);
                return new { Added = true, request.Expression,
                    WatchType = string.IsNullOrWhiteSpace(request.WatchType) ? "expression" : request.WatchType,
                    Context = new {
                    Project = shownProject, Module = shownModule, Procedure = shownProcedure },
                    Watches = watchState,
                    Verification = watches == IntPtr.Zero ? "Pending" : "ReadbackAvailable",
                    NextRead = watches == IntPtr.Zero ?
                        "Open the Watches pane and call debug_windows in a separate request to verify the expression and value." :
                        "Call debug_windows in a separate request to confirm the watch value after VBE refresh." };
            }
            finally
            {
                if (!completed)
                {
                    IntPtr remaining = native.Dialog("Ajouter un espion", "Add Watch");
                    if (remaining != IntPtr.Zero) native.Close(remaining);
                }
            }
        }

        /// <summary>Complète la modification de la surveillance sélectionnée.</summary>
        /// <param name="request">Requête avec l’expression et la cible attendues.</param>
        /// <returns>État vérifié de la surveillance après modification.</returns>
        public static object CompleteEditWatch(Request request)
        {
            return CompleteEditWatch(request, new NativeWatchProbe());
        }

        /// <summary>Remplit et vérifie Edit Watch avec la sonde indiquée.</summary>
        /// <param name="request">Expression et emplacement attendus.</param>
        /// <param name="native">Sonde des dialogues et du volet Watches.</param>
        /// <returns>État de la surveillance modifiée.</returns>
        internal static object CompleteEditWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Modifier un espion", "Edit Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Edit Watch dialog did not open.");
            bool completed = false;
            try
            {
                IntPtr edit = native.Item(dialog, 4853);
                IntPtr project = native.Item(dialog, 4858);
                IntPtr module = native.Item(dialog, 4857);
                IntPtr procedure = native.Item(dialog, 4856);
                IntPtr ok = native.Item(dialog, 1);
                if (edit == IntPtr.Zero || project == IntPtr.Zero || module == IntPtr.Zero || ok == IntPtr.Zero)
                    throw new InvalidOperationException("The native Edit Watch controls changed.");
                string original = native.Text(edit);
                string shownProject = native.Text(project);
                string shownModule = native.Text(module);
                string shownProcedure = procedure == IntPtr.Zero ? null : native.Text(procedure);
                string shownContext = shownModule + (string.IsNullOrWhiteSpace(shownProcedure) ? "" : "." + shownProcedure);
                if (!string.Equals(original, request.Expression, StringComparison.Ordinal) ||
                    !string.Equals(shownProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(shownContext, request.Context, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The native Edit Watch selection or context changed.");
                if (!string.IsNullOrWhiteSpace(request.WatchType))
                {
                    int typeId = request.WatchType == "break_when_true" ? 4851 :
                        request.WatchType == "break_when_changed" ? 4852 : 4850;
                    IntPtr typeButton = native.Item(dialog, typeId);
                    if (typeButton == IntPtr.Zero || !native.Click(typeButton))
                        throw new InvalidOperationException("The requested native watch type is unavailable.");
                    native.Pause(30);
                    if (!native.Checked(typeButton))
                        throw new InvalidOperationException("The requested native watch type was not selected.");
                }
                native.Replace(edit, request.NewExpression);
                if (!string.Equals(native.Text(edit), request.NewExpression, StringComparison.Ordinal))
                    throw new InvalidOperationException("The new watch expression was not reflected by the native edit control.");
                if (!native.Click(ok))
                    throw new InvalidOperationException("The Edit Watch dialog refused its OK command.");
                IntPtr error = IntPtr.Zero;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    native.Pause(50);
                    if (native.Dialog("Modifier un espion", "Edit Watch") == IntPtr.Zero)
                    { completed = true; break; }
                    error = native.Dialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
                    if (error != IntPtr.Zero) break;
                }
                if (error != IntPtr.Zero)
                {
                    string detail = native.Message(error);
                    native.Close(error);
                    throw new InvalidOperationException("VBE rejected the edited watch: " + detail);
                }
                if (!completed) throw new InvalidOperationException("The Edit Watch dialog did not close after OK.");
                IntPtr root = native.VbeRoot();
                IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                    native.Pane(native.Children(root), "Espions", "Watch", "Watches");
                if (pane == IntPtr.Zero)
                    return new { Edited = true, OldExpression = request.Expression, request.NewExpression,
                        request.Context, Verification = "Pending", VerificationPending = true,
                        Limit = "The Watches pane is not visible; edit cannot be read back." };
                bool verified = false;
                for (int attempt = 0; attempt < 20; attempt++)
                {
                    bool newPresent = native.WatchMatches(pane, request.NewExpression, request.Context) == 1;
                    bool oldAbsent = request.NewExpression == request.Expression ||
                        native.WatchMatches(pane, request.Expression, request.Context) == 0;
                    if (newPresent && oldAbsent) { verified = true; break; }
                    native.Pause(50);
                }
                return new { Edited = true, OldExpression = request.Expression, request.NewExpression,
                    request.Context, Verification = verified ? "ReadbackVerified" : "Pending",
                    VerificationPending = !verified,
                    Limit = "Expression readback does not independently prove a changed watch break condition." };
            }
            finally
            {
                if (!completed)
                {
                    IntPtr remaining = native.Dialog("Modifier un espion", "Edit Watch");
                    if (remaining != IntPtr.Zero) native.Close(remaining);
                }
            }
        }

        /// <summary>Évalue une expression dans Quick Watch, puis lit et ferme le dialogue natif.</summary>
        /// <param name="request">Expression et contexte de code attendus.</param>
        /// <returns>Valeur observée et état du dialogue.</returns>
        public static object CompleteQuickWatch(Request request)
        {
            return CompleteQuickWatch(request, new NativeWatchProbe());
        }

        /// <summary>Exécute Quick Watch avec une sonde injectable.</summary>
        /// <param name="request">Expression et contexte de code attendus.</param>
        /// <param name="native">Sonde du dialogue Quick Watch.</param>
        /// <returns>Valeur accessible retournée par le dialogue.</returns>
        internal static object CompleteQuickWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Espion express", "Quick Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Quick Watch dialog did not open.");
            try
            {
                IntPtr expressionControl = native.Item(dialog, 4751);
                IntPtr valueControl = native.Item(dialog, 4752);
                IntPtr contextControl = native.Item(dialog, 4753);
                if (expressionControl == IntPtr.Zero || valueControl == IntPtr.Zero ||
                    contextControl == IntPtr.Zero || native.Item(dialog, 2) == IntPtr.Zero)
                    throw new InvalidOperationException("The native Quick Watch controls changed.");
                string expression = native.Text(expressionControl);
                string value = native.Text(valueControl);
                string context = native.Text(contextControl);
                if (!string.Equals(expression, request.Expression, StringComparison.Ordinal) ||
                    !context.StartsWith(request.Project + "." + request.Module + ".", StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(request.Procedure) &&
                     !string.Equals(context, request.Project + "." + request.Module + "." + request.Procedure,
                         StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Quick Watch expression or context differs from the selected code.");
                return new { Expression = expression, Value = value, Context = context,
                    Verification = "NativeDialogReadback",
                    Limit = "Evaluating a VBA expression may call user code; a displayed value is valid for this paused context only." };
            }
            finally { native.Close(dialog); }
        }

        /// <summary>Sélectionne une ligne Watches correspondant à l’expression et au contexte requis.</summary>
        /// <param name="request">Expression et contexte à faire correspondre.</param>
        /// <returns>Expression sélectionnée et résultat de l’action.</returns>
        public static object SelectWatch(Request request)
        {
            return SelectWatch(request, new NativeWatchProbe());
        }

        /// <summary>Sélectionne la ligne avec la sonde native fournie.</summary>
        /// <param name="request">Expression et contexte à rechercher.</param>
        /// <param name="native">Sonde des lignes Watches.</param>
        /// <returns>Nombre de correspondances et résultat de sélection.</returns>
        internal static object SelectWatch(Request request, IWatchProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Expression) ||
                string.IsNullOrWhiteSpace(request.Context))
                throw new ArgumentException("Expression and Context are required.");
            IntPtr root = native.VbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                native.Pane(native.Children(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The native Watches pane must be visible.");
            int matches = native.WatchMatches(pane, request.Expression, request.Context);
            if (matches != 1)
                throw new InvalidOperationException("Expected one matching native watch; found " + matches + ".");
            if (!native.SelectWatchRow(pane, request.Expression, request.Context))
                throw new InvalidOperationException("The native watch row does not support selection.");
            return new { Selected = true, request.Expression, request.Context };
        }

        /// <summary>Vérifie que l’expression demandée n’apparaît plus dans le volet Watches.</summary>
        /// <param name="request">Expression et contexte à vérifier.</param>
        /// <returns>Nombre de correspondances restantes et état de vérification.</returns>
        public static object VerifyWatchRemoved(Request request)
        {
            return VerifyWatchRemoved(request, new NativeWatchProbe());
        }

        /// <summary>Vérifie la suppression d’une surveillance avec une sonde injectable.</summary>
        /// <param name="request">Expression et contexte concernés.</param>
        /// <param name="native">Sonde des volets Watches.</param>
        /// <returns>État observé après vérification.</returns>
        internal static object VerifyWatchRemoved(Request request, IWatchProbe native)
        {
            IntPtr root = native.VbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                native.Pane(native.Children(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero)
                return new { Removed = false, VerificationPending = true,
                    Error = "The Watches pane is no longer visible; absence cannot be verified." };
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.WatchMatches(pane, request.Expression, request.Context) == 0)
                    return new { Removed = true, VerificationPending = false, Error = (string)null };
                native.Pause(50);
            }
            return new { Removed = false, VerificationPending = true,
                Error = "The selected watch is still visible after the native command." };
        }

        /// <summary>Retourne les lignes UI Automation du volet Watches qui correspondent à l’expression et au contexte.</summary>
        /// <param name="pane">Handle du volet Watches.</param>
        /// <param name="expression">Expression recherchée.</param>
        /// <param name="context">Contexte parent facultatif.</param>
        /// <returns>Lignes accessibles correspondantes.</returns>
        private static List<AutomationElement> MatchingWatchRows(IntPtr pane, string expression, string context)
        {
            AutomationElement root = AutomationElement.FromHandle(pane);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
            AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, condition);
            var matches = new List<AutomationElement>();
            for (int index = 0; index < elements.Count; index++)
            {
                string raw = elements[index].Current.Name;
                Match parsed = Regex.Match(raw, @"^\s*(.*?)\s+(?:Valeur|Value)\s+.*?\s+Type\s+.*?\s+(?:Contexte|Context)\s+(.*?)\s*$",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (parsed.Success && string.Equals(parsed.Groups[1].Value, expression, StringComparison.Ordinal) &&
                    string.Equals(parsed.Groups[2].Value, context, StringComparison.OrdinalIgnoreCase))
                    matches.Add(elements[index]);
            }
            return matches;
        }

        /// <summary>Extrait le texte principal accessible d’un dialogue natif.</summary>
        /// <param name="dialog">Handle du dialogue.</param>
        /// <returns>Texte de message accessible, ou chaîne vide si aucun n’est trouvé.</returns>
        private static string AccessibleDialogMessage(IntPtr dialog)
        {
            AutomationElement root = AutomationElement.FromHandle(dialog);
            AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            for (int index = 0; index < elements.Count; index++)
            {
                if (elements[index].Current.ControlType != ControlType.Text) continue;
                string name = elements[index].Current.Name;
                if (!string.IsNullOrWhiteSpace(name) && name != "OK" && name != "Aide" && name != "Help") return name;
            }
            return "Native validation error.";
        }

        /// <summary>Ferme un dialogue natif et attend brièvement sa disparition.</summary>
        /// <param name="dialog">Handle du dialogue.</param>
        private static void CloseDialog(IntPtr dialog)
        {
            IntPtr cancel = GetDlgItem(dialog, 2);
            if (cancel != IntPtr.Zero) PostMessage(cancel, BmClick, IntPtr.Zero, IntPtr.Zero);
        }

        /// <summary>Recherche une fenêtre de dialogue visible dont le titre correspond à l’un des noms fournis.</summary>
        /// <param name="titles">Titres possibles, y compris leurs variantes localisées.</param>
        /// <returns>Handle du dialogue trouvé, ou zéro.</returns>
        private static IntPtr FindDialog(params string[] titles)
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid != currentPid || ClassName(handle) != "#32770" || !IsWindowVisible(handle)) return true;
                string title = WindowText(handle);
                foreach (string candidate in titles)
                    if (string.Equals(title, candidate, StringComparison.OrdinalIgnoreCase))
                    { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Recherche la fenêtre racine du VBE dans le processus hôte courant.</summary>
        /// <returns>Handle du VBE, ou zéro si aucune fenêtre ne correspond.</returns>
        private static IntPtr FindVbeRoot()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid == currentPid && ClassName(handle) == "wndclass_desked_gsk" && IsWindowVisible(handle))
                { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Énumère les fenêtres enfants visibles d’une fenêtre racine.</summary>
        /// <param name="root">Handle de la fenêtre dont les enfants sont énumérés.</param>
        /// <returns>Handles de fenêtres enfants visibles.</returns>
        private static List<IntPtr> ChildWindows(IntPtr root)
        {
            var result = new List<IntPtr>();
            EnumChildWindows(root, (handle, parameter) => {
                if (ClassName(handle) == "VbaWindow" && IsWindowVisible(handle)) result.Add(handle);
                return true;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Recherche le premier volet dont le titre correspond sans tenir compte de la casse.</summary>
        /// <param name="panes">Fenêtres candidates.</param>
        /// <param name="names">Titres possibles du volet.</param>
        /// <returns>Handle du premier volet correspondant, ou zéro.</returns>
        private static IntPtr FindPane(IEnumerable<IntPtr> panes, params string[] names)
        {
            foreach (IntPtr pane in panes)
            {
                string title = WindowText(pane);
                foreach (string name in names)
                    if (string.Equals(title, name, StringComparison.OrdinalIgnoreCase)) return pane;
            }
            return IntPtr.Zero;
        }

        /// <summary>Lit les éléments accessibles d’un contrôle de liste de débogage.</summary>
        /// <param name="handle">Handle du contrôle.</param>
        /// <returns>Éléments et valeurs lisibles du volet.</returns>
        private static object ReadList(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new { Available = false, Items = new object[0], Error = "Window is not visible.",
                Coverage = "UIAExposedRowsOnly" };
            try
            {
                AutomationElement root = AutomationElement.FromHandle(handle);
                var listCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
                AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, listCondition);
                var items = new List<object>();
                for (int index = 0; index < elements.Count; index++)
                {
                    AutomationElement element = elements[index];
                    string raw = element.Current.Name;
                    object pattern;
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                        raw = ((ValuePattern)pattern).Current.Value;
                    // Skip the native empty-list placeholder before querying its UIA ancestry.
                    if (ParseDebugRow(raw, null) == null) continue;
                    object parsed = ParseDebugRow(raw, ItemPath(element));
                    if (parsed != null) items.Add(parsed);
                }
                return new { Available = true, Items = items.ToArray(), Error = (string)null,
                    Coverage = "UIAExposedRowsOnly" };
            }
            catch (Exception ex) { return new { Available = true, Items = new object[0], Error = ex.Message,
                Coverage = "UIAExposedRowsOnly" }; }
        }

        /// <summary>Analyse une ligne affichée des volets Locals ou Watches et lui associe son chemin de parenté.</summary>
        /// <param name="raw">Texte brut de la ligne accessible.</param>
        /// <param name="itemPath">Segments du chemin UI Automation.</param>
        /// <returns>Valeurs découpées de la ligne et chemin d’élément.</returns>
        internal static object ParseDebugRow(string raw, string[] itemPath)
        {
            // MSForms' localized list rows flatten all columns into one accessible string.
            Match match = Regex.Match(raw ?? "", @"^Expression\s+(.*?)\s+Valeur\s+(.*?)\s+Type\s+(.*?)\s*$",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                match = Regex.Match(raw ?? "", @"^Expression\s+(.*?)\s+Value\s+(.*?)\s+Type\s+(.*?)\s*$",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match watch = Regex.Match(raw ?? "", @"^\s*(.*?)\s+Valeur\s+(.*?)\s+Type\s+(.*?)\s+Contexte\s+(.*?)\s*$",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!watch.Success)
                watch = Regex.Match(raw ?? "", @"^\s*(.*?)\s+Value\s+(.*?)\s+Type\s+(.*?)\s+Context\s+(.*?)\s*$",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (match.Success && string.IsNullOrWhiteSpace(match.Groups[1].Value) &&
                (match.Groups[2].Value.IndexOf("Aucune variable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 match.Groups[2].Value.IndexOf("No variables", StringComparison.OrdinalIgnoreCase) >= 0))
                return null;
            itemPath = itemPath ?? new string[0];
            return new { Raw = raw, Expression = match.Success ? match.Groups[1].Value :
                    watch.Success ? watch.Groups[1].Value : null,
                Value = match.Success ? match.Groups[2].Value.Trim() :
                    watch.Success ? watch.Groups[2].Value.Trim() : null,
                Type = match.Success ? match.Groups[3].Value :
                    watch.Success ? watch.Groups[3].Value : null,
                Context = watch.Success ? watch.Groups[4].Value : null,
                PathSegments = itemPath,
                Depth = Math.Max(0, itemPath.Length - 1),
                Parsed = match.Success || watch.Success };
        }

        /// <summary>Construit le chemin hiérarchique des noms accessibles entre la racine du volet et un élément.</summary>
        /// <param name="element">Élément UI Automation cible.</param>
        /// <returns>Segments de chemin dans l’ordre racine vers élément.</returns>
        private static string[] ItemPath(AutomationElement element)
        {
            var segments = new List<string>();
            AutomationElement current = element;
            for (int depth = 0; depth < 16 && current != null &&
                current.Current.ControlType == ControlType.ListItem; depth++)
            {
                string segment = ParseItemPathSegment(current.Current.Name);
                if (segment == null) break;
                segments.Add(segment);
                current = TreeWalker.RawViewWalker.GetParent(current);
            }
            segments.Reverse();
            return segments.ToArray();
        }

        /// <summary>Normalise un segment de chemin d’élément en retirant son marqueur d’expansion éventuel.</summary>
        /// <param name="name">Nom accessible brut.</param>
        /// <returns>Nom utilisable comme segment.</returns>
        internal static string ParseItemPathSegment(string name)
        {
            Match match = Regex.Match(name ?? "",
                @"^Expression\s+(.*?)\s+(?:Valeur|Value)\s+", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                match = Regex.Match(name ?? "",
                    @"^\s*(.*?)\s+(?:Valeur|Value)\s+.*?\s+Type\s+.*?\s+(?:Contexte|Context)\s+",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        /// <summary>Détermine la procédure racine associée à une ligne Watches.</summary>
        /// <param name="element">Ligne UI Automation.</param>
        /// <returns>Contexte de procédure, ou chaîne vide s’il n’est pas présent.</returns>
        private static string WatchRootContext(AutomationElement element)
        {
            AutomationElement root = element;
            AutomationElement parent;
            while ((parent = TreeWalker.RawViewWalker.GetParent(root)) != null &&
                parent.Current.ControlType == ControlType.ListItem)
                root = parent;
            return ParseWatchContext(root.Current.Name);
        }

        /// <summary>Extrait le contexte de procédure affiché entre parenthèses dans une ligne Watch.</summary>
        /// <param name="name">Nom accessible de la ligne.</param>
        /// <returns>Contexte extrait, ou chaîne vide.</returns>
        internal static string ParseWatchContext(string name)
        {
            Match match = Regex.Match(name ?? "",
                @"\s+(?:Contexte|Context)\s+(.*?)\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        /// <summary>Lit le texte de l’éditeur et les informations d’accessibilité du volet Immediate.</summary>
        /// <param name="handle">Handle du volet.</param>
        /// <returns>Instantané du contenu Immediate.</returns>
        private static object ReadImmediate(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new { Available = false, Text = (string)null, Error = "Window is not visible." };
            try
            {
                return new { Available = true, Text = ImmediateText(handle), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Text = (string)null, Error = ex.Message }; }
        }

        /// <summary>Recherche l’élément de document UI Automation dans le volet Immediate.</summary>
        /// <param name="handle">Handle du volet Immediate.</param>
        /// <returns>Élément de document textuel.</returns>
        /// <exception cref="InvalidOperationException">Aucun document accessible n’est disponible.</exception>
        private static AutomationElement ImmediateDocument(IntPtr handle)
        {
            AutomationElement root = AutomationElement.FromHandle(handle);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
            AutomationElementCollection documents = root.FindAll(TreeScope.Descendants, condition);
            if (documents.Count != 1)
                throw new InvalidOperationException("Expected one Immediate document; found " + documents.Count + ".");
            if (!documents[0].TryGetCurrentPattern(TextPattern.Pattern, out _))
                throw new InvalidOperationException("The Immediate document does not expose TextPattern.");
            return documents[0];
        }

        /// <summary>Lit le texte complet de l’élément document du volet Immediate.</summary>
        /// <param name="handle">Handle du volet.</param>
        /// <returns>Texte courant de l’éditeur Immediate.</returns>
        private static string ImmediateText(IntPtr handle)
        {
            var pattern = (TextPattern)ImmediateDocument(handle).GetCurrentPattern(TextPattern.Pattern);
            return pattern.DocumentRange.GetText(-1);
        }

        /// <summary>Ouvre le dialogue Call Stack depuis Locals, lit ses trames et ferme le dialogue.</summary>
        /// <param name="locals">Handle du volet Locals.</param>
        /// <returns>Trames lues et disponibilité du dialogue.</returns>
        private static object ReadCallStack(IntPtr locals)
        {
            IntPtr dialog = FindCallStackDialog();
            bool openedHere = dialog == IntPtr.Zero;
            if (openedHere)
            {
                if (locals == IntPtr.Zero)
                    return new { Available = false, Frames = new string[0], Error = "Locals window must be visible to open Call Stack without a shortcut." };
                IntPtr button = IntPtr.Zero;
                EnumChildWindows(locals, (handle, parameter) => {
                    if (ClassName(handle) == "Button" && GetDlgCtrlID(handle) == 4601 && IsWindowVisible(handle))
                    { button = handle; return false; }
                    return true;
                }, IntPtr.Zero);
                if (button == IntPtr.Zero || !PostMessage(button, BmClick, IntPtr.Zero, IntPtr.Zero))
                    return new { Available = false, Frames = new string[0], Error = "The native Call Stack button was not available." };
                for (int attempt = 0; attempt < 30 && dialog == IntPtr.Zero; attempt++)
                { PauseNative(50); dialog = FindCallStackDialog(); }
            }
            if (dialog == IntPtr.Zero)
                return new { Available = false, Frames = new string[0], Error = "Call Stack dialog did not open." };
            try
            {
                IntPtr list = IntPtr.Zero;
                EnumChildWindows(dialog, (handle, parameter) => {
                    if (ClassName(handle) == "ListBox") { list = handle; return false; }
                    return true;
                }, IntPtr.Zero);
                if (list == IntPtr.Zero) throw new InvalidOperationException("Call Stack list was not found.");
                object accessible;
                Guid iid = IidAccessible;
                int hr = AccessibleObjectFromWindow(list, ObjidClient, ref iid, out accessible);
                if (hr != 0 || !(accessible is Accessibility.IAccessible))
                    throw new COMException("Call Stack MSAA object was unavailable.", hr);
                var listAccess = (Accessibility.IAccessible)accessible;
                var frames = new List<string>();
                for (int index = 1; index <= Math.Min(listAccess.accChildCount, 256); index++)
                    frames.Add(listAccess.get_accName(index));
                return new { Available = true, Frames = frames.ToArray(), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Frames = new string[0], Error = ex.Message }; }
            finally
            {
                if (openedHere)
                {
                    EnumChildWindows(dialog, (handle, parameter) => {
                        if (ClassName(handle) == "Button" && GetDlgCtrlID(handle) == 2)
                        { PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); return false; }
                        return true;
                    }, IntPtr.Zero);
                }
            }
        }

        /// <summary>Recherche le dialogue natif Call Stack parmi les titres localisés connus.</summary>
        /// <returns>Handle du dialogue, ou zéro s’il n’est pas visible.</returns>
        private static IntPtr FindCallStackDialog()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                string title = WindowText(handle);
                if (pid == currentPid && ClassName(handle) == "#32770" && IsWindowVisible(handle) &&
                    (title == "Pile des appels" || title == "Call Stack"))
                { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Lit le nom de classe Win32 d’une fenêtre.</summary>
        /// <param name="handle">Handle à interroger.</param>
        /// <returns>Nom de classe, ou chaîne vide si la lecture échoue.</returns>
        private static string ClassName(IntPtr handle)
        { var text = new StringBuilder(128); GetClassName(handle, text, text.Capacity); return text.ToString(); }
        /// <summary>Lit le texte ou titre Win32 d’une fenêtre.</summary>
        /// <param name="handle">Handle à interroger.</param>
        /// <returns>Texte observé, ou chaîne vide si la lecture échoue.</returns>
        private static string WindowText(IntPtr handle)
        { var text = new StringBuilder(512); GetWindowText(handle, text, text.Capacity); return text.ToString(); }
    }
}
