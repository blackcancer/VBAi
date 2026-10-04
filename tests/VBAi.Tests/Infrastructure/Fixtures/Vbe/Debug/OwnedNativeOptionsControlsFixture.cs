namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Automation;
    using System.Windows.Forms;
    using VBAi;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        [DllImport("user32.dll", EntryPoint = "SetParent")]
        private static extern IntPtr OptionsFixtureSetParent(IntPtr child, IntPtr parent);

        private sealed class ParentLostOnComboSelection : NativeWindow, IDisposable
        {
            internal bool Armed;
            internal IntPtr PreviousParent;
            private int previousStyle;
            internal ParentLostOnComboSelection(IntPtr handle) { AssignHandle(handle); }
            protected override void WndProc(ref Message message)
            {
                bool detachAfterSelection = Armed && message.Msg == 0x14E;
                base.WndProc(ref message);
                if (detachAfterSelection)
                {
                    Armed = false;
                    previousStyle = OptionsFixtureGetStyle(Handle, -16);
                    PreviousParent = OptionsFixtureSetParent(Handle, IntPtr.Zero);
                    OptionsFixtureSetStyle(Handle, -16, previousStyle & ~unchecked((int)0x40000000));
                }
            }
            public void Dispose()
            {
                if (PreviousParent != IntPtr.Zero)
                {
                    OptionsFixtureSetStyle(Handle, -16, previousStyle);
                    OptionsFixtureSetParent(Handle, PreviousParent);
                }
                ReleaseHandle();
            }
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "CreateWindowExW")]
        private static extern IntPtr OptionsFixtureCreate(uint extended, string kind, string text, uint style,
            int x, int y, int width, int height, IntPtr parent, IntPtr identifier, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr OptionsFixtureText(IntPtr window, int message, IntPtr argument, string text);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr OptionsFixtureInteger(IntPtr window, int message, IntPtr argument, IntPtr value);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr OptionsFixtureReadText(IntPtr window, int message, IntPtr capacity, StringBuilder text);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int OptionsFixtureGetStyle(IntPtr window, int index);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern int OptionsFixtureSetStyle(IntPtr window, int index, int value);
        [DllImport("user32.dll", EntryPoint = "EnableWindow")]
        private static extern bool OptionsFixtureEnable(IntPtr window, bool enabled);

        private static string OptionsFixtureReadEditText(IntPtr window)
        {
            int length = OptionsFixtureInteger(window, 0x000E, IntPtr.Zero, IntPtr.Zero).ToInt32();
            if (length < 0 || length > 4096) throw new InvalidOperationException("Owned native combo edit length is invalid.");
            var value = new StringBuilder(length + 1);
            int copied = OptionsFixtureReadText(window, 0x000D, new IntPtr(value.Capacity), value).ToInt32();
            if (copied != length) throw new InvalidOperationException("Owned native combo edit changed during read.");
            return value.ToString();
        }

        /// <summary>Invoque une méthode privée pour tester ses vrais contrôles natifs, en propageant son exception réelle.</summary>
        private static object InvokeOptionsMethod(object instance, string name, params object[] arguments)
        {
            try
            {
                var type = instance == null ? typeof(VbeDebugWindows) : instance.GetType();
                return type.GetMethod(name, BindingFlags.NonPublic | (instance == null ? BindingFlags.Static : BindingFlags.Instance)).Invoke(instance, arguments);
            }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }

        /// <summary>Héberge de vrais ComboBox/ListBox standards dans un vrai dialogue #32770 détenu; seul le propriétaire des couleurs est simulé.</summary>
        private sealed class OwnedNativeOptionsControls : IDisposable
        {
            internal readonly AutomationNode Root = new AutomationNode { Name = "Options", Kind = ControlType.Window };
            internal readonly AutomationNode Tab = new AutomationNode { Name = "Editor Format", Kind = ControlType.TabItem };
            internal readonly AutomationNode Categories = new AutomationNode { Name = "Code Colors", Kind = ControlType.List };
            internal readonly List<AutomationNode> CategoryItems = new List<AutomationNode>();
            internal readonly List<AutomationNode> Palettes = new List<AutomationNode>();
            internal readonly Dictionary<string, int[]> Colours = new Dictionary<string, int[]> {
                ["Normal"] = new[] { 0, 0, 0 }, ["Comment"] = new[] { 1, 2, 0 }, ["Keyword"] = new[] { 2, 1, 2 } };
            internal readonly AutomationHost Host;
            internal readonly List<Tuple<int, int, IntPtr>> Notifications = new List<Tuple<int, int, IntPtr>>();
            internal IntPtr Font, Size, List;
            internal string CurrentCategory = "Normal";
            internal Action<string> OnCategoryNotification;
            internal Action<int, int, IntPtr> OnControlNotification;

            internal OwnedNativeOptionsControls()
            {
                Root.Add(Tab.With(SelectionItemPattern.Pattern));
                Root.Add(Categories.With(SelectionPattern.Pattern));
                foreach (string category in new[] { "Normal", "Comment", "Keyword" })
                    CategoryItems.Add(Categories.Add(new AutomationNode { Name = category, Selected = category == "Normal" }.With(SelectionItemPattern.Pattern)));
                foreach (string name in new[] { "Foreground", "Background", "Indicator" })
                    Palettes.Add(Root.Add(new AutomationNode { Name = name, Kind = ControlType.ComboBox }));
                Host = new AutomationHost(Root, optionsDialog: true, configure: owner =>
                {
                    owner.NativeMessage = ObserveOwnerNotification;
                    List = Create(owner.Handle, "ListBox", 4905, 0x50010001);
                    Categories.NativeHandle = List.ToInt32();
                    for (int i = 0; i < CategoryItems.Count; i++)
                    {
                        int index = i;
                        OptionsFixtureText(List, 0x180, IntPtr.Zero, CategoryItems[i].Name);
                        CategoryItems[i].SelectedAction = () =>
                        {
                            foreach (var item in CategoryItems) item.Selected = ReferenceEquals(item, CategoryItems[index]);
                            OptionsFixtureInteger(List, 0x186, new IntPtr(index), IntPtr.Zero);
                        };
                    }
                    OptionsFixtureInteger(List, 0x186, IntPtr.Zero, IntPtr.Zero);
                    for (int i = 0; i < Palettes.Count; i++)
                    {
                        IntPtr combo = Create(owner.Handle, "ComboBox", 501 + i, 0x50210203);
                        Palettes[i].NativeHandle = combo.ToInt32();
                        foreach (string label in new[] { "Automatic", "", "" }) OptionsFixtureText(combo, 0x143, IntPtr.Zero, label);
                        OptionsFixtureInteger(combo, 0x14e, IntPtr.Zero, IntPtr.Zero);
                    }
                    Font = Create(owner.Handle, "ComboBox", 510, 0x50210202);
                    foreach (string label in new[] { "Consolas", "Courier New", "Duplicate", "Duplicate" }) OptionsFixtureText(Font, 0x143, IntPtr.Zero, label);
                    OptionsFixtureText(Font, 0x000C, IntPtr.Zero, "Consolas");
                    Size = Create(owner.Handle, "ComboBox", 511, 0x50210202);
                    OptionsFixtureText(Size, 0x000C, IntPtr.Zero, "10");
                }, noActivate: true);
            }

            /// <summary>Crée uniquement une fenêtre enfant standard du dialogue détenu.</summary>
            private static IntPtr Create(IntPtr parent, string kind, int identifier, uint style)
            {
                int x = identifier == 510 || identifier == 511 ? 205 : 5;
                int y = identifier == 4905 ? 5 : identifier >= 501 && identifier <= 503
                    ? 140 + (identifier - 501) * 30 : identifier == 510 ? 5 : 130;
                var window = OptionsFixtureCreate(0, kind, "", style, x, y, 180, 120, parent, new IntPtr(identifier), IntPtr.Zero, IntPtr.Zero);
                if (window == IntPtr.Zero) throw new InvalidOperationException("Owned native options control creation failed.");
                return window;
            }

            /// <summary>Applique les palettes simulées uniquement lorsque la véritable notification Win32 atteint le parent détenu.</summary>
            private void ObserveOwnerNotification(Message message)
            {
                if (message.Msg != 0x111) return;
                int identifier = (int)(message.WParam.ToInt64() & 0xffff), code = (int)(message.WParam.ToInt64() >> 16);
                Notifications.Add(Tuple.Create(identifier, code, message.LParam));
                OnControlNotification?.Invoke(identifier, code, message.LParam);
                if (identifier == 4905 && code == 1 && message.LParam == List)
                {
                    int selected = OptionsFixtureInteger(List, 0x188, IntPtr.Zero, IntPtr.Zero).ToInt32();
                    if (selected < 0 || selected >= CategoryItems.Count) return;
                    CurrentCategory = CategoryItems[selected].Name;
                    for (int i = 0; i < Palettes.Count; i++)
                        OptionsFixtureInteger(new IntPtr(Palettes[i].NativeHandle.Value), 0x14e, new IntPtr(Colours[CurrentCategory][i]), IntPtr.Zero);
                    Palettes[2].Enabled = CurrentCategory != "Keyword";
                    OptionsFixtureEnable(new IntPtr(Palettes[2].NativeHandle.Value), Palettes[2].Enabled);
                    OnCategoryNotification?.Invoke(CurrentCategory);
                }
                else if (identifier >= 501 && identifier <= 503 && code == 1)
                    Colours[CurrentCategory][identifier - 501] = OptionsFixtureInteger(message.LParam, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();
            }
            public void Dispose() { Host.Dispose(); }
        }
    }
}
