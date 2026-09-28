using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    /// <summary>Routines natives pour modifier et mesurer l’affichage du code dans une fenêtre VBE.</summary>
internal static partial class VbeDebugWindows
    {
        /// <summary>Rectangle Win32 obtenu pour une fenêtre ou un contrôle natif.</summary>
[StructLayout(LayoutKind.Sequential)] internal struct ViewRect { /// <summary>Stores the left,top,right,bottom used by ViewRect.</summary>
public int Left, Top, Right, Bottom; }
        /// <summary>Lit les limites d’une fenêtre Win32.</summary>
        /// <param name="hwnd">Handle de la fenêtre à mesurer.</param>
        /// <param name="rect">Rectangle écran retourné par Windows.</param>
        /// <returns><see langword="true"/> si le rectangle a été obtenu.</returns>
[DllImport("user32.dll", EntryPoint = "GetWindowRect")] private static extern bool NativeViewBounds(IntPtr hwnd, out ViewRect rect);
        /// <summary>Frontière injectable de lecture des limites d’une fenêtre.</summary>
        /// <param name="hwnd">Handle de la fenêtre.</param>
        /// <param name="rect">Rectangle écran résultant.</param>
        /// <returns>Indique si la lecture a réussi.</returns>
internal delegate bool ViewBoundsReader(IntPtr hwnd, out ViewRect rect);
        /// <summary>Reads native window geometry; isolated in deterministic window tests.</summary>
        internal static ViewBoundsReader ViewBounds = NativeViewBounds;

        /// <summary>Clique sur le bouton natif de vue procédure ou module de la fenêtre de code reconnue.</summary>
        /// <param name="caption">Légende exacte de la fenêtre VBE à cibler.</param>
        /// <param name="procedure"><see langword="true"/> pour la vue Procédure; <see langword="false"/> pour Module.</param>
        /// <returns>Méthode et coordonnées locales du bouton activé.</returns>
internal static object ChangeCodeView(string caption, bool procedure)
        {
            IntPtr root = FindVbeRoot();
            var windows = ChildWindows(root).Where(window => WindowText(window) == caption).ToArray();
            if (root == IntPtr.Zero || windows.Length != 1) throw new InvalidOperationException("The exact native code window was not identified.");
            IntPtr windowHandle = windows[0];
            for (IntPtr ancestor = windowHandle; ancestor != IntPtr.Zero; ancestor = ObjectBrowserParent(ancestor))
                if (!ObjectBrowserEnabled(ancestor)) throw new InvalidOperationException("The code window is disabled by a modal window.");
            var bars = new List<ViewRect>();
            var buttons = new List<Tuple<IntPtr, ViewRect>>();
            EnumChildWindows(windowHandle, (child, ignored) =>
            {
                if (!IsWindowVisible(child) || !ViewBounds(child, out ViewRect bounds)) return true;
                string kind = ClassName(child);
                if (kind == "ScrollBar" && bounds.Right-bounds.Left > bounds.Bottom-bounds.Top) bars.Add(bounds);
                if (kind == "ObtbarWndClass") buttons.Add(Tuple.Create(child,bounds));
                return true;
            }, IntPtr.Zero);
            if (bars.Count != 1) throw new InvalidOperationException("The code window must expose exactly one shared horizontal scrollbar.");
            ViewRect bar = bars[0];
            var matches = buttons.Where(item => item.Item2.Right == bar.Left && item.Item2.Top == bar.Top && item.Item2.Bottom == bar.Bottom).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The native procedure/module view toolbar was not uniquely identified.");
            var target = matches[0];
            int width = target.Item2.Right-target.Item2.Left, height = target.Item2.Bottom-target.Item2.Top;
            if (height < 10 || height > 64 || width < height*1.5 || width > height*3 || !ObjectBrowserEnabled(target.Item1))
                throw new InvalidOperationException("Unrecognized native view-button geometry or disabled control.");
            int x = procedure ? width/4 : 3*width/4, y = height/2;
            IntPtr point = new IntPtr((y << 16) | x);
            // These owner-drawn buttons expose no UIA/MSAA action. Send mouse
            // messages only to their verified HWND using its local dimensions.
            SendMessageInt(target.Item1, 0x0201, new IntPtr(1), point);
            SendMessageInt(target.Item1, 0x0202, IntPtr.Zero, point);
            return new { Method = "Native view toolbar, local button center", Width = width, Height = height, X = x, Y = y };
        }
    }
}
