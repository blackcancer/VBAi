using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Web.Script.Serialization;
using System.Drawing;

public static class VbePaletteInspection
{
    private delegate bool EnumProc(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int length);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);
    [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr window);
    [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
    [DllImport("gdi32.dll")] private static extern uint GetPixel(IntPtr dc, int x, int y);
    private static Thread worker;
    public static string Result;
    private static int operation;
    private static string recoveryPath;
    private static List<PaletteRow> original;
    public sealed class PaletteRow
    {
        public string Name;
        public int Foreground, Background, Indicator;
    }

    public static void BeginPreview(uint processId, string path)
    {
        operation = 1;
        recoveryPath = path;
        Begin(processId);
    }

    public static void BeginRestore(uint processId)
    {
        if (original == null) throw new InvalidOperationException("No original palette was captured.");
        operation = 2;
        Begin(processId);
    }

    public static void BeginVerifyRestore(uint processId) { operation = 3; Begin(processId); }

    public static void Begin(uint processId)
    {
        Result = null;
        worker = new Thread(() => Inspect(processId)) { IsBackground = true };
        worker.SetApartmentState(ApartmentState.MTA);
        worker.Start();
    }

    public static bool Wait() { return worker.Join(10000); }

    private static IntPtr Find(uint processId)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((window, parameter) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            var text = new StringBuilder(256);
            GetWindowText(window, text, text.Capacity);
            if (owner == processId && text.ToString() == "Options") { found = window; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static void Inspect(uint processId)
    {
        IntPtr dialog = IntPtr.Zero;
        bool accept = false;
        try
        {
            for (int attempt = 0; attempt < 100 && dialog == IntPtr.Zero; attempt++) { Thread.Sleep(100); dialog = Find(processId); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Options dialog was not found.");
            var root = AutomationElement.FromHandle(dialog);
            var tabs = root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
            bool selected = false;
            foreach (AutomationElement tab in tabs)
                if (tab.Current.Name.IndexOf("Format", StringComparison.OrdinalIgnoreCase) >= 0)
                { ((SelectionItemPattern)tab.GetCurrentPattern(SelectionItemPattern.Pattern)).Select(); selected = true; break; }
            if (!selected) throw new InvalidOperationException("The formatting tab was not found.");
            Thread.Sleep(150);
            var rows = new List<object>();
            foreach (AutomationElement element in root.FindAll(TreeScope.Descendants, Condition.TrueCondition))
            {
                if (element.Current.IsOffscreen) continue;
                string className = element.Current.ClassName;
                IntPtr window = new IntPtr(element.Current.NativeWindowHandle);
                var patterns = new List<string>();
                foreach (var pattern in element.GetSupportedPatterns()) patterns.Add(pattern.ProgrammaticName);
                var data = new List<long>();
                int count = -1, selection = -1;
                if (className == "ComboBox")
                {
                    count = SendMessage(window, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32();
                    selection = SendMessage(window, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();
                    for (int i = 0; i < Math.Min(count, 32); i++) data.Add(SendMessage(window, 0x150, new IntPtr(i), IntPtr.Zero).ToInt64());
                }
                rows.Add(new { Name = element.Current.Name, Class = className, Id = element.Current.AutomationId,
                    Type = element.Current.ControlType.ProgrammaticName, Patterns = patterns, Count = count, Selection = selection, ItemData = data });
            }
            var categories = new List<PaletteRow>();
            var list = root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "4912"));
            IntPtr foregroundHandle = Control(root, "4913"), backgroundHandle = Control(root, "4914"), indicatorHandle = Control(root, "4935");
            foreach (AutomationElement item in list.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)))
            {
                ((SelectionItemPattern)item.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
                Thread.Sleep(75);
                categories.Add(new PaletteRow { Name = item.Current.Name,
                    Foreground = SendMessage(foregroundHandle, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32(),
                    Background = SendMessage(backgroundHandle, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32(),
                    Indicator = SendMessage(indicatorHandle, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32() });
            }
            if (operation != 0)
            {
                if (categories.Count != 10) throw new InvalidOperationException("Unexpected palette category count.");
                if (operation == 3)
                {
                    for (int i = 0; i < categories.Count; i++)
                        if (categories[i].Name != original[i].Name || categories[i].Foreground != original[i].Foreground ||
                            categories[i].Background != original[i].Background || categories[i].Indicator != original[i].Indicator)
                            throw new InvalidOperationException("The restored palette differs from its initial values.");
                    Result = "{\"RestoreVerified\":true}";
                    return;
                }
                if (operation == 1)
                {
                    original = categories;
                    System.IO.File.WriteAllText(recoveryPath, new JavaScriptSerializer().Serialize(original));
                }
                // Approximate Visual Studio Community Dark with the native VBE palette.
                // Its fixed colors cannot reproduce VS's RGB values. Dark cyan is more
                // readable on black than native blue; muted green avoids neon comments.
                // Order: normal, selected, syntax error, current statement, breakpoint,
                // comment, keyword, identifier, bookmark, call return.
                int[] foregroundIndices = { 8, 16, 13, 1, 16, 3, 4, 8, 8, 1 };
                int[] backgroundIndices = { 1, 2, 1, 15, 5, 1, 1, 1, 1, 15 };
                int index = 0;
                foreach (AutomationElement item in list.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)))
                {
                    ((SelectionItemPattern)item.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
                    Thread.Sleep(75);
                    if (operation == 2 && item.Current.Name != original[index].Name) throw new InvalidOperationException("Palette categories changed before restore.");
                    SetColor(foregroundHandle, 4913, operation == 1 ? foregroundIndices[index] : original[index].Foreground);
                    SetColor(backgroundHandle, 4914, operation == 1 ? backgroundIndices[index] : original[index].Background);
                    index++;
                }
                Result = new JavaScriptSerializer().Serialize(new { Applied = operation == 1, Restored = operation == 2, Categories = categories });
                accept = true;
                return;
            }
            var foreground = root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "4913"));
            ((ExpandCollapsePattern)foreground.GetCurrentPattern(ExpandCollapsePattern.Pattern)).Expand();
            Thread.Sleep(100);
            var colors = new List<string>();
            foreach (AutomationElement item in foreground.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)))
                colors.Add(item.Current.Name);
            ((ExpandCollapsePattern)foreground.GetCurrentPattern(ExpandCollapsePattern.Pattern)).Collapse();
            var swatches = new List<object>();
            IntPtr combo = new IntPtr(foreground.Current.NativeWindowHandle);
            int width = (int)foreground.Current.BoundingRectangle.Width, height = (int)foreground.Current.BoundingRectangle.Height;
            for (int i = 0; i < 17; i++)
            {
                SendMessage(combo, 0x14e, new IntPtr(i), IntPtr.Zero);
                Thread.Sleep(25);
                using (var bitmap = new Bitmap(width, height))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        IntPtr dc = graphics.GetHdc();
                        try { PrintWindow(combo, dc, 0); } finally { graphics.ReleaseHdc(dc); }
                    }
                    IntPtr live = GetDC(combo);
                    uint color;
                    try { color = GetPixel(live, width / 3, height / 2); }
                    finally { ReleaseDC(combo, live); }
                    swatches.Add(new { Index = i, Color = color.ToString("X8") });
                }
            }
            Result = new JavaScriptSerializer().Serialize(new { Controls = rows, Categories = categories, Colors = colors, Swatches = swatches });
        }
        catch (Exception error) { Result = new JavaScriptSerializer().Serialize(new { Error = error.ToString() }); }
        finally
        {
            if (dialog != IntPtr.Zero)
            {
                IntPtr cancel = GetDlgItem(dialog, accept ? 1 : 2);
                if (cancel != IntPtr.Zero) PostMessage(cancel, 0xf5, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }

    private static IntPtr Control(AutomationElement root, string id)
    {
        var element = root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
        if (element == null || element.Current.NativeWindowHandle == 0) throw new InvalidOperationException("Missing native color control " + id);
        return new IntPtr(element.Current.NativeWindowHandle);
    }

    private static void SetColor(IntPtr combo, int id, int index)
    {
        if (combo == IntPtr.Zero || SendMessage(combo, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32() != 17)
            throw new InvalidOperationException("Unexpected native color selector.");
        if (SendMessage(combo, 0x14e, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
            throw new InvalidOperationException("The native color selection was rejected.");
        SendMessage(GetParent(combo), 0x111, new IntPtr((1 << 16) | id), combo);
        if (SendMessage(combo, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32() != index)
            throw new InvalidOperationException("The native color selection was not retained.");
    }
}
