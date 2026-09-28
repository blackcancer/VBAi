using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    internal static partial class VbeDebugWindows
    {
        /// <summary>Nœud MSAA en lecture seule, injectable sans activer le concepteur.</summary>
        internal interface IToolboxAccessibleNode : IDisposable
        {
            /// <summary>Rôle du nœud lui-même (0) ou du childID exact.</summary>
            int Role(int child);
            /// <summary>Flags natifs complets sans action.</summary>
            int State(int child);
            /// <summary>Libellé observé ; ce texte ne sert pas à identifier le nœud.</summary>
            string Name(int child);
            /// <summary>Nombre de descendants directs exposés par le fournisseur.</summary>
            int Count { get; }
            /// <summary>Interface enfant ou null pour un enfant simple ; chaque interface est libérée par le lecteur.</summary>
            IToolboxAccessibleNode Child(int child);
        }

        /// <summary>Racine et descendants bornés du fournisseur MSAA natif.</summary>
        private sealed class ToolboxAccessibleNode : IToolboxAccessibleNode
        {
            private readonly Accessibility.IAccessible accessible;
            internal ToolboxAccessibleNode(Accessibility.IAccessible value) { accessible = value; }
            public int Role(int child) => Convert.ToInt32(accessible.get_accRole(child), CultureInfo.InvariantCulture);
            public int State(int child) => Convert.ToInt32(accessible.get_accState(child), CultureInfo.InvariantCulture);
            public string Name(int child) => accessible.get_accName(child);
            public int Count => accessible.accChildCount;
            public IToolboxAccessibleNode Child(int child)
            {
                object value = accessible.get_accChild(child);
                if (value is Accessibility.IAccessible node) return new ToolboxAccessibleNode(node);
                if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                return null;
            }
            public void Dispose() { if (Marshal.IsComObject(accessible)) Marshal.ReleaseComObject(accessible); }
        }

        /// <summary>Ouvre seulement une interface MSAA ; aucune méthode d'action n'est exposée.</summary>
        private static IToolboxAccessibleNode OpenToolboxAccessibleNode(IntPtr handle)
        {
            object value;
            Guid iid = IidAccessible;
            int hr = AccessibleObjectFromWindow(handle, ObjidClient, ref iid, out value);
            if (hr != 0 || !(value is Accessibility.IAccessible node))
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                throw new InvalidOperationException("The native Toolbox MSAA provider is unavailable.");
            }
            return new ToolboxAccessibleNode(node);
        }

        /// <summary>Lit un unique serveur groupant les pages, exclusivement dans le processus du VBE.</summary>
        private static NavigationSurface ReadNativeToolboxPages(IntPtr window, uint owner)
        {
            var result = new NavigationSurface { Provider = "MSAA", Caption = WindowText(window),
                Coverage = "Observed Toolbox pages only. Buttons are not exposed by this fallback; page actions are not qualified.",
                ButtonsExposed = false };
            try
            {
                GetWindowThreadProcessId(window, out uint actual);
                if (owner == 0 || actual != owner || !IsWindowVisible(window) ||
                    !(ClassName(window) == "VbaWindow" || ClassName(window).StartsWith("F3 MinFrame ", StringComparison.Ordinal)) ||
                    (result.Caption != "Toolbox" && result.Caption != "Boîte à outils")) throw new InvalidOperationException();
                var servers = new HashSet<IntPtr>();
                EnumChildWindows(window, (h, p) => {
                    GetWindowThreadProcessId(h, out uint pid);
                    if (pid == owner && IsWindowVisible(h) && ClassName(h).StartsWith("F3 Server ", StringComparison.Ordinal)) servers.Add(h);
                    return true;
                }, IntPtr.Zero);
                if (servers.Count == 0 || servers.Count > 64) throw new InvalidOperationException();
                var groups = new List<IntPtr>();
                foreach (IntPtr server in servers)
                    using (var node = OpenToolboxAccessibleNode(server))
                        if (node.Role(0) == 20) groups.Add(server);
                if (groups.Count != 1) throw new InvalidOperationException();
                IntPtr handle = groups[0];
                using (var node = OpenToolboxAccessibleNode(handle))
                    result = ReadToolboxPages(window, handle, result.Caption, node);
                GetWindowThreadProcessId(window, out actual);
                GetWindowThreadProcessId(handle, out uint serverOwner);
                if (actual != owner || serverOwner != owner || !IsWindowVisible(window) || !IsWindowVisible(handle) || WindowText(window) != result.Caption ||
                    !ClassName(handle).StartsWith("F3 Server ", StringComparison.Ordinal))
                    throw new InvalidOperationException();
            }
            catch (Exception)
            {
                result.Available = false; result.Nodes = new NavigationNode[0]; result.Identity = null;
                result.Error = "The native Toolbox pages are absent, ambiguous, inaccessible or changed during readback.";
            }
            return result;
        }

        /// <summary>Collecte exactement une liste de pages, sans déduire l'absence des boutons de palette.</summary>
        internal static NavigationSurface ReadToolboxPages(IntPtr window, IntPtr server, string caption, IToolboxAccessibleNode root)
        {
            var state = new NavigationSurface { Caption = caption, Provider = "MSAA", ButtonsExposed = false,
                Coverage = "Observed Toolbox pages only. Buttons are not exposed by this fallback; page actions are not qualified.",
                Identity = "msaa:" + window.ToInt64().ToString(CultureInfo.InvariantCulture) + ":" + server.ToInt64().ToString(CultureInfo.InvariantCulture) };
            if (window == IntPtr.Zero || server == IntPtr.Zero || root == null || root.Role(0) != 20)
                throw new InvalidOperationException("The exact Toolbox grouping provider is required.");
            var pages = new List<NavigationNode>();
            int visited = 0, lists = 0;
            int? listState = null;
            ReadToolboxPageBranch(root, state.Identity, "0", 0, ref visited, ref lists, ref listState, pages);
            if (lists != 1 || pages.Count == 0 || pages.Count > 128)
                throw new InvalidOperationException("The native Toolbox page list is absent, ambiguous or exceeds its bound.");
            state.MSAAContainerState = listState;
            state.Nodes = pages.ToArray(); state.Available = true;
            return state;
        }

        /// <summary>Parcourt les childID en ordre natif, sans coordonnées ni noms utilisés comme identités.</summary>
        private static void ReadToolboxPageBranch(IToolboxAccessibleNode node, string identity, string path, int depth,
            ref int visited, ref int lists, ref int? listState, List<NavigationNode> pages)
        {
            if (depth > 8 || ++visited > 512) throw new InvalidOperationException("The Toolbox accessibility tree exceeds its bound.");
            int count = node.Count;
            if (count < 0 || count > 128) throw new InvalidOperationException("The Toolbox accessibility child count is invalid.");
            if (node.Role(0) == 60)
            {
                if (++lists > 1) throw new InvalidOperationException("Multiple Toolbox page lists are ambiguous.");
                listState = node.State(0);
                if (listState < 0) throw new InvalidOperationException("The Toolbox page list state is unreadable.");
                for (int child = 1; child <= count; child++)
                {
                    if (++visited > 512 || node.Role(child) != 37) throw new InvalidOperationException("The Toolbox page list is unreadable.");
                    string name = node.Name(child);
                    int flags = node.State(child);
                    if (string.IsNullOrEmpty(name) || name.Length > 256 || flags < 0) throw new InvalidOperationException("The Toolbox page state is unreadable.");
                    pages.Add(new NavigationNode { Token = identity + ":" + path + "." + child.ToString(CultureInfo.InvariantCulture),
                        ParentToken = identity + ":" + path, Name = name, Kind = "MSAA.PageTab", Enabled = (flags & 1) == 0,
                        Selected = null, ObservedSelected = (flags & 2) != 0, MSAAState = flags,
                        ActionUnavailableReason = "Toolbox page selection is observed only; changing pages through MSAA is not qualified." });
                }
                return;
            }
            for (int child = 1; child <= count; child++)
                using (var branch = node.Child(child))
                    if (branch != null) ReadToolboxPageBranch(branch, identity, path + "." + child.ToString(CultureInfo.InvariantCulture),
                        depth + 1, ref visited, ref lists, ref listState, pages);
        }
    }
}