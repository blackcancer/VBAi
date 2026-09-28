using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Routes code-module double clicks in the native project tree to Monaco.</summary>
    internal sealed class EditorProjectNavigation : IDisposable
    {
        /// <summary>Instance VBE dont les projets et composants sont parcourus.</summary>
        private readonly object vbe;
        /// <summary>Contrôle WinForms utilisé pour distribuer l’ouverture sur le thread UI.</summary>
        private readonly Control dispatcher;
        /// <summary>Action qui ouvre le module sélectionné dans l’éditeur moderne.</summary>
        private readonly Action<IEditorModule> open;
        /// <summary>Hooks attachés aux arborescences de projets actuellement ouvertes.</summary>
        private readonly List<TreeHook> hooks = new List<TreeHook>();
        /// <summary>Minuterie qui découvre périodiquement les fenêtres d’explorateur de projets.</summary>
        private readonly Timer timer = new Timer { Interval = 1000 };
        /// <summary>Crée les hooks de navigation et démarre leur actualisation périodique.</summary>
        /// <param name="vbe">Instance d’automatisation VBE.</param>
        /// <param name="dispatcher">Contrôle servant à distribuer l’ouverture sur le thread UI.</param>
        /// <param name="open">Action appelée avec l’adaptateur du composant sélectionné.</param>
        internal EditorProjectNavigation(object vbe, Control dispatcher, Action<IEditorModule> open)
        {
            this.vbe = vbe; this.dispatcher = dispatcher; this.open = open;
            timer.Tick += (s, e) => Refresh(); Refresh(); timer.Start();
        }
        /// <summary>Repère les arborescences de projets nouvelles et retire les handles détruits.</summary>
        private void Refresh()
        {
            hooks.RemoveAll(h => h.Handle == IntPtr.Zero);
            try
            {
                foreach (dynamic window in ((dynamic)vbe).Windows)
                {
                    if ((int)window.Type != 6) continue; // vbext_wt_ProjectWindow
                    foreach (IntPtr handle in FindProjectTrees(System.Diagnostics.Process.GetCurrentProcess().Id, (string)window.Caption))
                    {
                        uint process;
                        if (GetWindowThreadProcessId(handle, out process) == GetCurrentThreadId() &&
                            !hooks.Any(h => h.Handle == handle)) hooks.Add(new TreeHook(handle, Route));
                    }
                }
            }
            catch (Exception error) { LoadLog.Write("Monaco project navigation: " + error.GetType().Name); }
        }
        /// <summary>Ouvre le composant de code sélectionné lorsque le double clic vise son nœud.</summary>
        /// <returns><see langword="true"/> lorsqu’une ouverture a été mise en file d’attente.</returns>
        private bool Route()
        {
            try
            {
                dynamic component = ((dynamic)vbe).SelectedVBComponent;
                if (component == null || !IsCodeComponent((int)component.Type)) return false;
                object project = component.Collection.Parent;
                var module = new EditorVbeModule(vbe, project, component);
                dispatcher.BeginInvoke(new Action(() => open(module)));
                return true;
            }
            catch (Exception error) { LoadLog.Write("Monaco navigation refused: " + error.GetType().Name); return false; }
        }
        /// <summary>Indique si le type VBComponent désigne un module, une classe ou un module de document.</summary>
        /// <param name="type">Valeur numérique de VBComponent.Type.</param>
        /// <returns><see langword="true"/> pour les types 1, 2 ou 100.</returns>
        internal static bool IsCodeComponent(int type) => type == 1 || type == 2 || type == 100;
        /// <summary>Recherche les contrôles TreeView de l’explorateur de projets dans les fenêtres du processus hôte.</summary>
        /// <param name="processId">Identifiant du processus qui héberge VBE.</param>
        /// <param name="caption">Titre de la fenêtre de projet à retrouver.</param>
        /// <returns>Handles uniques des arbres de projet correspondant au titre.</returns>
        internal static IntPtr[] FindProjectTrees(int processId, string caption)
        {
            var result = new List<IntPtr>();
            EnumWindows((root, ignored) =>
            {
                GetWindowThreadProcessId(root, out uint pid);
                var kind = new StringBuilder(128); GetClassName(root, kind, kind.Capacity);
                if (pid != processId || kind.ToString() != "wndclass_desked_gsk") return true;
                EnumChildWindows(root, (pane, unused) =>
                {
                    var title = new StringBuilder(512); GetWindowText(pane, title, title.Capacity);
                    if (title.ToString() != caption) return true;
                    EnumChildWindows(pane, (child, value) =>
                    {
                        var name = new StringBuilder(128); GetClassName(child, name, name.Capacity);
                        if (name.ToString() == "SysTreeView32" && !result.Contains(child)) result.Add(child);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return result.ToArray();
        }
        /// <summary>Arrête la découverte et détache les hooks des fenêtres encore ouvertes.</summary>
        public void Dispose()
        { timer.Stop(); timer.Dispose(); foreach (var hook in hooks) if (hook.Handle != IntPtr.Zero) hook.ReleaseHandle(); hooks.Clear(); }
        /// <summary>Intercepte les doubles clics du TreeView natif d’un explorateur de projets.</summary>
        private sealed class TreeHook : NativeWindow
        {
            /// <summary>Route le nœud de module ciblé vers l’éditeur moderne.</summary>
            private readonly Func<bool> route;
            /// <summary>Associe le hook au TreeView d’un projet.</summary>
            /// <param name="handle">Handle de l’arborescence du projet.</param>
            /// <param name="route">Fonction qui ouvre le composant sélectionné.</param>
            internal TreeHook(IntPtr handle, Func<bool> route) { this.route = route; AssignHandle(handle); }
            /// <summary>Intercepte le double clic sans détourner les autres gestes du TreeView.</summary>
            /// <param name="message">Message Windows reçu par l’arborescence.</param>
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x0203) // WM_LBUTTONDBLCLK, delivered only to this native project tree
                {
                    long point = message.LParam.ToInt64();
                    var hit = new HitTest { X = unchecked((short)point), Y = unchecked((short)(point >> 16)) };
                    SendMessage(Handle, 0x1111, IntPtr.Zero, ref hit); // TVM_HITTEST
                    // Leave expansion glyphs, empty space and noncomponent nodes to the tree.
                    if ((hit.Flags & 0x46) != 0 && hit.Item != IntPtr.Zero &&
                        hit.Item == SendMessage(Handle, 0x110A, new IntPtr(9), IntPtr.Zero) && route())
                    {
                        // Keep VBE's backing code window alive. The queued Monaco activation
                        // runs after the native double-click has finished opening it.
                        base.WndProc(ref message); return;
                    }
                }
                base.WndProc(ref message);
            }
        }
        /// <summary>Coordonnées et nœud renvoyés par le message TreeView TVM_HITTEST.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct HitTest { /// <summary>Coordonnée horizontale du point testé.</summary>
public int X, Y; /// <summary>Indicateurs de la zone du nœud touchée.</summary>
public uint Flags; /// <summary>Handle du nœud sous le point testé.</summary>
public IntPtr Item; }
        /// <summary>Rappel Win32 appelé pour chaque fenêtre parcourue.</summary>
        /// <param name="handle">Fenêtre visitée.</param>
        /// <param name="parameter">Valeur opaque transmise au parcours.</param>
        /// <returns><see langword="true"/> pour continuer l’énumération.</returns>
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);
        /// <summary>Énumère les fenêtres de premier niveau.</summary>
        /// <param name="callback">Rappel appelé pour chaque fenêtre.</param>
        /// <param name="parameter">Valeur transmise au rappel.</param>
        /// <returns>Indique si l’énumération a pu se poursuivre.</returns>
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindow callback, IntPtr parameter);
        /// <summary>Lit le titre Unicode d’une fenêtre dans le tampon fourni.</summary>
        /// <param name="handle">Fenêtre dont le titre est lu.</param>
        /// <param name="value">Tampon de réception.</param>
        /// <param name="count">Capacité du tampon.</param>
        /// <returns>Nombre de caractères copiés.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, StringBuilder value, int count);
        /// <summary>Énumère les fenêtres enfants d’un parent.</summary>
        /// <param name="parent">Fenêtre parente.</param>
        /// <param name="callback">Rappel appelé pour chaque enfant.</param>
        /// <param name="parameter">Valeur transmise au rappel.</param>
        /// <returns>Indique si l’énumération a pu se poursuivre.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
        /// <summary>Lit le nom de classe Unicode d’une fenêtre.</summary>
        /// <param name="handle">Fenêtre inspectée.</param>
        /// <param name="value">Tampon de réception.</param>
        /// <param name="count">Capacité du tampon.</param>
        /// <returns>Nombre de caractères copiés.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder value, int count);
        /// <summary>Récupère les identifiants du thread et du processus propriétaires d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre.</param>
        /// <param name="process">Reçoit l’identifiant du processus.</param>
        /// <returns>Identifiant du thread propriétaire.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);
        /// <summary>Obtient l’identifiant du thread Windows courant.</summary>
        /// <returns>Identifiant du thread.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        /// <summary>Envoie au TreeView le message qui identifie l’élément situé sous un point.</summary>
        /// <param name="handle">Handle de l’arborescence.</param>
        /// <param name="message">Identifiant du message Windows.</param>
        /// <param name="wParam">Premier paramètre du message.</param>
        /// <param name="hit">Coordonnées fournies et informations de l’élément trouvé.</param>
        /// <returns>Résultat du message.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, ref HitTest hit);
        /// <summary>Envoie un message Windows avec deux paramètres pointeurs.</summary>
        /// <param name="handle">Fenêtre cible.</param>
        /// <param name="message">Identifiant du message.</param>
        /// <param name="wParam">Premier paramètre du message.</param>
        /// <param name="lParam">Second paramètre du message.</param>
        /// <returns>Valeur renvoyée par la procédure de fenêtre.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
    }
}
