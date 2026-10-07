using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Observe les fenêtres de formulaires VBA visibles dans le processus hôte.</summary>
    internal static partial class VbeDebugWindows
    {

        /// <summary>Lit la fenêtre propriétaire d’un handle sans accéder aux objets formulaire VBA.</summary>
        /// <param name="window">Handle dont le propriétaire doit être obtenu.</param>
        /// <param name="command">Commande native GetWindow, généralement GW_OWNER.</param>
        /// <returns>Handle de la fenêtre propriétaire, ou zéro.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindow")]
        private static extern IntPtr NativeRuntimeWindowOwner(IntPtr window, uint command);

        /// <summary>Reads the native owner without accessing a running VBA form instance.</summary>
        internal static Func<IntPtr, uint, IntPtr> RuntimeWindowOwner = NativeRuntimeWindowOwner;

        /// <summary>Énumère les fenêtres ThunderDFrame visibles appartenant à l’hôte courant.</summary>
        /// <returns>Fenêtres, erreurs d’énumération et limites de cette observation native.</returns>
        internal static object ReadRuntimeForms()
        {
            uint host = (uint)Process.GetCurrentProcess().Id;
            var windows = new List<object>();
            var errors = new List<string>();
            int detected = 0;
            bool complete = EnumWindows((window, ignored) =>
            {
                try
                {
                    GetWindowThreadProcessId(window, out uint owner);
                    if (owner != host || !IsWindowVisible(window) || ClassName(window) != "ThunderDFrame") return true;
                    detected++;
                    if (windows.Count >= 64) return true;
                    IntPtr parent = RuntimeWindowOwner(window, 4); // GW_OWNER
                    bool boundsRead = ViewBounds(window, out ViewRect bounds);
                    windows.Add(new
                    {
                        Handle = window.ToInt64(),
                        Caption = WindowText(window),
                        ClassName = "ThunderDFrame",
                        Visible = true,
                        Enabled = ObjectBrowserEnabled(window),
                        OwnerHandle = parent.ToInt64(),
                        OwnerEnabled = parent == IntPtr.Zero ? (bool?)null : ObjectBrowserEnabled(parent),
                        Bounds = boundsRead ? (object)new { bounds.Left, bounds.Top, Width = bounds.Right - bounds.Left, Height = bounds.Bottom - bounds.Top } : null,
                        Project = (string)null,
                        ProjectIdentityVerified = false
                    });
                }
                catch (Exception error) { errors.Add(error.Message); }
                return true;
            }, IntPtr.Zero);
            return new
            {
                HostProcessId = host,
                Windows = windows,
                DetectedCount = detected,
                Truncated = detected > windows.Count,
                EnumerationSucceeded = complete,
                Errors = errors,
                Source = "Visible top-level ThunderDFrame windows in this host process",
                Limit = "Native window observation only: does not resolve a form instance to a VBProject or prove initialization. Hidden forms and other window classes are excluded. Handles can expire; no controls or values are read. Mode 2 does not imply this list is empty."
            };
        }
    }
}
