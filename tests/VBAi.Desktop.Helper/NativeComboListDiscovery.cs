using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

namespace VBAi.Desktop.Helper
{
    /// <summary>Read-only exact dropdown-list discovery for an already expanded native ComboBox.</summary>
    public static class NativeComboListDiscovery
    {
        [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct ComboBoxInfo
        {
            public uint Size;
            public Rect Item, Button;
            public uint ButtonState;
            public IntPtr Combo, ItemHandle, List;
        }
        [DllImport("user32.dll", SetLastError = true)] private static extern bool GetComboBoxInfo(IntPtr combo, ref ComboBoxInfo info);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder value, int size);

        private static IntPtr NativeAncestor(AutomationElement item)
        {
            for (int depth = 0; item != null && depth < 32; depth++, item = TreeWalker.RawViewWalker.GetParent(item))
                if (item.Current.NativeWindowHandle != 0)
                    return new IntPtr(unchecked((long)(uint)item.Current.NativeWindowHandle));
            return IntPtr.Zero;
        }

        public static bool MatchesNativeList(int comboPid, uint comboTid, int listPid, uint listTid,
            int expectedPid, uint expectedTid, string listClass, bool listVisible,
            bool comboDesktopMatches, bool listDesktopMatches)
            => expectedPid > 0 && expectedTid != 0 && comboPid == expectedPid && listPid == expectedPid &&
               comboTid == expectedTid && listTid == expectedTid && listVisible &&
               comboDesktopMatches && listDesktopMatches &&
               (listClass == "ComboLBox" || listClass != null && listClass.Contains("LISTBOX"));

        private static AutomationElement[] RawNodes(AutomationElement root)
        {
            var found = new List<AutomationElement>();
            var pending = new Queue<AutomationElement>(); pending.Enqueue(root);
            while (pending.Count != 0)
            {
                var next = pending.Dequeue();
                if (found.Count >= 128) throw new InvalidOperationException("Native combo UIA tree bound exceeded.");
                found.Add(next);
                for (var child = TreeWalker.RawViewWalker.GetFirstChild(next); child != null;
                    child = TreeWalker.RawViewWalker.GetNextSibling(child))
                {
                    if (found.Count + pending.Count >= 128)
                        throw new InvalidOperationException("Native combo UIA tree bound exceeded.");
                    pending.Enqueue(child);
                }
            }
            return found.ToArray();
        }

        /// <summary>Returns a unique virtual ListItem under the native dropdown HWND; performs no selection.</summary>
        public static AutomationElement RequireExactItem(IntPtr combo, string expectedName, int expectedPid,
            uint expectedTid, string expectedDesktop, Func<uint, string> desktopName,
            Action<object> observation)
        {
            if (combo == IntPtr.Zero || string.IsNullOrEmpty(expectedName) || expectedPid <= 0 || expectedTid == 0 ||
                string.IsNullOrEmpty(expectedDesktop) || desktopName == null || observation == null)
                throw new InvalidOperationException("Exact native combo discovery requires an owned identity and desktop.");
            var clock = Stopwatch.StartNew(); object last = null;
            while (clock.Elapsed < TimeSpan.FromSeconds(3))
            {
                var info = new ComboBoxInfo { Size = (uint)Marshal.SizeOf(typeof(ComboBoxInfo)) };
                if (!GetComboBoxInfo(combo, ref info) || info.Combo != combo || info.List == IntPtr.Zero ||
                    !IsWindow(combo) || !IsWindow(info.List))
                    throw new InvalidOperationException("Native combo/list HWND relation unavailable after expansion.");
                uint comboPid, listPid;
                uint comboTid = GetWindowThreadProcessId(combo, out comboPid);
                uint listTid = GetWindowThreadProcessId(info.List, out listPid);
                var cls = new StringBuilder(128); GetClassName(info.List, cls, cls.Capacity);
                string nativeClass = cls.ToString();
                bool visible = IsWindowVisible(info.List);
                bool comboDesktop = string.Equals(desktopName(comboTid), expectedDesktop, StringComparison.OrdinalIgnoreCase);
                bool listDesktop = string.Equals(desktopName(listTid), expectedDesktop, StringComparison.OrdinalIgnoreCase);
                if (!MatchesNativeList((int)comboPid, comboTid, (int)listPid, listTid,
                    expectedPid, expectedTid, nativeClass, visible, comboDesktop, listDesktop))
                {
                    observation(new { Phase = "NativeComboListDiscoveryRefused", ComboHandle = combo.ToInt64(),
                        ListHandle = info.List.ToInt64(), ListClass = nativeClass,
                        ComboPid = comboPid, ComboTid = comboTid, ListPid = listPid, ListTid = listTid,
                        ListVisible = visible, ComboDesktopMatches = comboDesktop, ListDesktopMatches = listDesktop,
                        ActionDelivered = false });
                    throw new InvalidOperationException("Native combo/list owner, class or private desktop changed.");
                }
                var root = AutomationElement.FromHandle(info.List);
                var nodes = RawNodes(root);
                int exactNameCount = 0;
                var matches = new List<AutomationElement>();
                foreach (var node in nodes)
                {
                    if (!string.Equals(node.Current.Name, expectedName, StringComparison.Ordinal)) continue;
                    exactNameCount++;
                    if (node.Current.ControlType == ControlType.ListItem && node.Current.ProcessId == expectedPid &&
                        node.Current.IsEnabled && !node.Current.IsOffscreen &&
                        unchecked((uint)NativeAncestor(node).ToInt64()) == unchecked((uint)info.List.ToInt64()))
                        matches.Add(node);
                }
                last = new { ComboHandle = combo.ToInt64(), ListHandle = info.List.ToInt64(),
                    ListClass = nativeClass, ComboPid = comboPid, ComboTid = comboTid,
                    ListPid = listPid, ListTid = listTid, ListVisible = visible,
                    RawNodeCount = nodes.Length, ExactNameCount = exactNameCount,
                    EligibleListItemCount = matches.Count };
                if (matches.Count > 1) break;
                if (visible && matches.Count == 1)
                {
                    observation(new { Phase = "ExactNativeComboListObserved", State = last });
                    return matches[0];
                }
                Thread.Sleep(50); // Provider settlement only; no mutation or selection retry.
            }
            observation(new { Phase = "NativeComboListDiscoveryRefused", State = last });
            throw new InvalidOperationException("The exact native combo dropdown item remained absent or ambiguous.");
        }
    }
}
