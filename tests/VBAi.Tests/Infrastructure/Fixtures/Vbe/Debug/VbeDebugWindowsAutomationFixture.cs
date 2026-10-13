namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows;
    using System.Windows.Automation;
    using System.Windows.Automation.Provider;
    using System.Windows.Automation.Text;
    using System.Windows.Forms;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        /// <summary>Preserves the low 32 bits of a Windows handle or WM_GETOBJECT identifier without checked narrowing.</summary>
        /// <param name="value">Native value, whose object identifiers can be zero-extended or sign-extended on x64.</param>
        /// <returns>The same low 32 bits represented as the signed integer required by UI Automation.</returns>
        internal static int NativeProtocolInt32(IntPtr value) => unchecked((int)value.ToInt64());

        [ComVisible(true)]
        public sealed class AutomationNode : IRawElementProviderSimple, IRawElementProviderFragmentRoot,
            ISelectionItemProvider, ISelectionProvider, IValueProvider, IToggleProvider,
            IRangeValueProvider, IExpandCollapseProvider, ITextProvider
        {
            private static int nextId;
            private readonly int id = Interlocked.Increment(ref nextId);
            public AutomationNode Parent;
            public readonly List<AutomationNode> Children = new List<AutomationNode>();
            public readonly HashSet<int> Patterns = new HashSet<int>();
            public string Name;
            public ControlType Kind = ControlType.ListItem;
            public string Text = "";
            public bool Enabled = true;
            public bool Offscreen;
            public bool Password;
            public bool Selected;
            public bool Expanded = true;
            public bool HideCollapsedChildren;
            public bool FailName;
            public bool FailValue;
            public bool FailSelection;
            public int? RuntimeIdentity;
            public bool ReadOnly;
            public int FocusCount;
            public int SelectionCount;
            public ToggleState ToggleState;
            public double Number = 25;
            public IntPtr Window;
            public int? NativeHandle;
            /// <summary>PID réellement propriétaire du provider, avec surcharge explicite pour les scénarios d'identité refusée.</summary>
            public int? ProcessId;
            public Action SelectedAction;
            public Action<int> OnPropertyRead;
            public Action<int> OnPatternRead;

            public AutomationNode Add(AutomationNode child) { child.Parent = this; Children.Add(child); return child; }
            public AutomationNode With(params AutomationPattern[] patterns) { foreach (var pattern in patterns) Patterns.Add(pattern.Id); return this; }
            private AutomationNode Root { get { var root = this; while (root.Parent != null) root = root.Parent; return root; } }
            public ProviderOptions ProviderOptions => ProviderOptions.ServerSideProvider;
            public IRawElementProviderSimple HostRawElementProvider => Parent == null && Window != IntPtr.Zero
                ? AutomationInteropProvider.HostProviderFromHandle(Window) : null;
            public object GetPatternProvider(int patternId) { OnPatternRead?.Invoke(patternId); return Patterns.Contains(patternId) ? this : null; }
            public object GetPropertyValue(int propertyId)
            {
                OnPropertyRead?.Invoke(propertyId);
                if (propertyId == AutomationElementIdentifiers.NameProperty.Id) { if (FailName) throw new ElementNotAvailableException(); return Name; }
                if (propertyId == AutomationElementIdentifiers.ControlTypeProperty.Id) return Kind.Id;
                if (propertyId == AutomationElementIdentifiers.IsEnabledProperty.Id) return Enabled;
                if (propertyId == AutomationElementIdentifiers.IsOffscreenProperty.Id) return Offscreen;
                if (propertyId == AutomationElementIdentifiers.IsPasswordProperty.Id) return Password;
                if (propertyId == AutomationElementIdentifiers.IsControlElementProperty.Id || propertyId == AutomationElementIdentifiers.IsContentElementProperty.Id) return true;
                if (propertyId == AutomationElementIdentifiers.IsKeyboardFocusableProperty.Id) return true;
                if (propertyId == AutomationElementIdentifiers.NativeWindowHandleProperty.Id) return NativeHandle ?? (Parent == null ? NativeProtocolInt32(Window) : 0);
                if (propertyId == AutomationElementIdentifiers.ProcessIdProperty.Id) return ProcessId ?? System.Diagnostics.Process.GetCurrentProcess().Id;
                return null;
            }
            public IRawElementProviderFragment Navigate(NavigateDirection direction)
            {
                var shown = HideCollapsedChildren && !Expanded ? new AutomationNode[0] : Children.ToArray();
                if (direction == NavigateDirection.Parent) return Parent;
                if (direction == NavigateDirection.FirstChild) return shown.FirstOrDefault();
                if (direction == NavigateDirection.LastChild) return shown.LastOrDefault();
                if (Parent == null) return null;
                int index = Parent.Children.IndexOf(this) + (direction == NavigateDirection.NextSibling ? 1 : -1);
                return index >= 0 && index < Parent.Children.Count ? Parent.Children[index] : null;
            }
            public int[] GetRuntimeId() { return new[] { AutomationInteropProvider.AppendRuntimeId, RuntimeIdentity ?? id }; }
            public Rect BoundingRectangle => new Rect(0, 0, 100, 100);
            public IRawElementProviderSimple[] GetEmbeddedFragmentRoots() { return null; }
            public void SetFocus() { FocusCount++; }
            public IRawElementProviderFragmentRoot FragmentRoot => Root;
            public IRawElementProviderFragment ElementProviderFromPoint(double x, double y) { return this; }
            public IRawElementProviderFragment GetFocus() { return this; }
            bool ISelectionItemProvider.IsSelected { get { if (FailSelection) throw new InvalidOperationException("selection unreadable"); return Selected; } }
            IRawElementProviderSimple ISelectionItemProvider.SelectionContainer => Parent;
            public void Select() { Selected = true; SelectionCount++; SelectedAction?.Invoke(); }
            public void AddToSelection() { Select(); }
            public void RemoveFromSelection() { Selected = false; }
            public IRawElementProviderSimple[] GetSelection() { return Children.Where(child => child.Selected).Cast<IRawElementProviderSimple>().ToArray(); }
            public bool CanSelectMultiple => false;
            public bool IsSelectionRequired => false;
            string IValueProvider.Value { get { if (FailValue) throw new InvalidOperationException("value unreadable"); return Text; } }
            bool IValueProvider.IsReadOnly => ReadOnly;
            public void SetValue(string value) { Text = value; }
            public void Toggle() { ToggleState = ToggleState == ToggleState.On ? ToggleState.Off : ToggleState.On; }
            ToggleState IToggleProvider.ToggleState { get { if (FailValue) throw new InvalidOperationException("toggle unreadable"); return ToggleState; } }
            double IRangeValueProvider.Value => Number;
            bool IRangeValueProvider.IsReadOnly => false;
            double IRangeValueProvider.Minimum => 0;
            double IRangeValueProvider.Maximum => 100;
            double IRangeValueProvider.LargeChange => 10;
            double IRangeValueProvider.SmallChange => 1;
            void IRangeValueProvider.SetValue(double value) { Number = value; }
            public void Expand() { Expanded = true; }
            public void Collapse() { Expanded = false; }
            public ExpandCollapseState ExpandCollapseState => Expanded ? ExpandCollapseState.Expanded : ExpandCollapseState.Collapsed;
            public ITextRangeProvider DocumentRange => new AutomationRange(this);
            public SupportedTextSelection SupportedTextSelection => SupportedTextSelection.Single;
            ITextRangeProvider[] ITextProvider.GetSelection() { return new[] { DocumentRange }; }
            public ITextRangeProvider[] GetVisibleRanges() { return new[] { DocumentRange }; }
            public ITextRangeProvider RangeFromChild(IRawElementProviderSimple childElement) { return DocumentRange; }
            public ITextRangeProvider RangeFromPoint(Point screenLocation) { return DocumentRange; }
        }

        [ComVisible(true)]
        public sealed class AutomationRange : ITextRangeProvider
        {
            private readonly AutomationNode node;
            public AutomationRange(AutomationNode node) { this.node = node; }
            public ITextRangeProvider Clone() { return new AutomationRange(node); }
            public bool Compare(ITextRangeProvider range) { return range is AutomationRange other && ReferenceEquals(other.node, node); }
            public int CompareEndpoints(TextPatternRangeEndpoint endpoint, ITextRangeProvider range, TextPatternRangeEndpoint targetEndpoint) { return 0; }
            public void ExpandToEnclosingUnit(TextUnit unit) { }
            public ITextRangeProvider FindAttribute(int attributeId, object value, bool backward) { return null; }
            public ITextRangeProvider FindText(string text, bool backward, bool ignoreCase) { return null; }
            public object GetAttributeValue(int attributeId) { return AutomationElement.NotSupported; }
            public double[] GetBoundingRectangles() { return new double[0]; }
            public IRawElementProviderSimple GetEnclosingElement() { return node; }
            public string GetText(int maxLength) { return maxLength < 0 ? node.Text : node.Text.Substring(0, Math.Min(maxLength, node.Text.Length)); }
            public int Move(TextUnit unit, int count) { return 0; }
            public void MoveEndpointByRange(TextPatternRangeEndpoint endpoint, ITextRangeProvider range, TextPatternRangeEndpoint targetEndpoint) { }
            public int MoveEndpointByUnit(TextPatternRangeEndpoint endpoint, TextUnit unit, int count) { return 0; }
            public void Select() { node.SelectionCount++; }
            public void AddToSelection() { Select(); }
            public void RemoveFromSelection() { }
            public void ScrollIntoView(bool alignToTop) { }
            public IRawElementProviderSimple[] GetChildren() { return new IRawElementProviderSimple[0]; }
        }

        internal sealed class AutomationHost : IDisposable
        {
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "CreateWindowExW")]
            private static extern IntPtr CreateOwnedDialog(uint extended, string kind, string text, uint style,
                int x, int y, int width, int height, IntPtr parent, IntPtr identifier, IntPtr instance, IntPtr parameter);
            [DllImport("user32.dll", EntryPoint = "DestroyWindow")]
            private static extern bool DestroyOwnedDialog(IntPtr window);
            private readonly Thread thread;
            private readonly ManualResetEvent ready = new ManualResetEvent(false);
            private ProviderForm form;
            private ProviderDialog dialog;
            private Control dispatcher;
            private Exception startupError;
            private Exception teardownError;
            private int disposeEntered;
            // Read from the waiting test thread only to describe the original deadline.
            private string teardownStage = "not requested";
            private int attemptedProviderDisconnects;
            private int completedProviderDisconnects;
            internal int DisconnectedProviderCount { get; private set; }
            internal bool OwnerThreadExited => !thread.IsAlive;
            [DllImport("UIAutomationCore.dll", PreserveSig = true)]
            private static extern int UiaDisconnectProvider([MarshalAs(UnmanagedType.Interface)] IRawElementProviderSimple provider);
            [DllImport("UIAutomationCore.dll", EntryPoint = "UiaReturnRawElementProvider")]
            private static extern IntPtr ReleaseWindowProviderMap(IntPtr window, IntPtr wParam, IntPtr lParam, IntPtr provider);
            public readonly AutomationNode Root;
            public IntPtr Handle;
            public AutomationHost(AutomationNode root, bool optionsDialog = false, Action<AutomationHost> configure = null, bool noActivate = false)
            {
                Root = root;
                thread = new Thread(() =>
                {
                    try
                    {
                        if (optionsDialog)
                        {
                            dispatcher = new Control();
                            if (dispatcher.Handle == IntPtr.Zero) throw new InvalidOperationException("Owned UIA dispatcher creation failed.");
                            Handle = CreateOwnedDialog(noActivate ? 0x08000080u : 0x80u, "#32770", "VBAi isolated Options fixture", 0x10cf0000,
                                20, 20, 400, 300, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                            if (Handle == IntPtr.Zero) throw new InvalidOperationException("Owned native #32770 dialog creation failed.");
                            root.Window = Handle;
                            dialog = new ProviderDialog(root, Handle);
                            configure?.Invoke(this);
                            ready.Set();
                            System.Windows.Forms.Application.Run(new ApplicationContext());
                        }
                        else
                        {
                            form = new ProviderForm { Provider = root, Text = "VBAi isolated UIA fixture", ShowInTaskbar = false, NoActivate = noActivate };
                            form.Shown += (sender, args) => { Handle = form.Handle; root.Window = Handle; configure?.Invoke(this); ready.Set(); };
                            System.Windows.Forms.Application.Run(form);
                        }
                    }
                    catch (Exception error) { startupError = error; ready.Set(); }
                })
                { IsBackground = true };
                thread.SetApartmentState(ApartmentState.STA); thread.Start();
                if (!ready.WaitOne(TimeSpan.FromSeconds(10))) throw new TimeoutException("Isolated UIA provider did not start.");
                if (startupError != null) throw new InvalidOperationException("Isolated UIA provider failed.", startupError);
            }
            public void Dispose()
            {
                if (Interlocked.Exchange(ref disposeEntered, 1) != 0) return;
                var teardownClock = System.Diagnostics.Stopwatch.StartNew();
                Volatile.Write(ref teardownStage, "owner dispatch pending");
                (dispatcher ?? (Control)form).BeginInvoke(new Action(() =>
                {
                    // Disconnect outside synchronous native-message handling. Reentrant WM_GETOBJECT
                    // must not publish the same provider while UI Automation releases it.
                    if (dialog != null) dialog.Disconnecting = true;
                    else form.Disconnecting = true;
                    try
                    {
                        // Destroy the owned HWND before disconnecting its providers, as required
                        // by the UI Automation provider lifecycle. No provider is republished.
                        Volatile.Write(ref teardownStage, "window closure");
                        if (dialog != null)
                        {
                            dialog.Release();
                            if (!DestroyOwnedDialog(Handle))
                                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                        }
                        else form.Close();
                    }
                    catch (Exception error) { teardownError = error; }
                    finally
                    {
                        int? simulatedNativeHandle = Root.NativeHandle;
                        try
                        {
                            // Negative scenarios hide the HWND from the application under test.
                            // During teardown only, UIA needs the actual owned root identity to
                            // resolve its runtime ID and disconnect the existing provider.
                            Root.NativeHandle = null;
                            Volatile.Write(ref teardownStage, "window provider-map release");
                            ReleaseWindowProviderMap(Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                            Volatile.Write(ref teardownStage, "provider disconnection");
                            var pending = new Stack<AutomationNode>();
                            var visited = new HashSet<AutomationNode>();
                            pending.Push(Root);
                            while (pending.Count != 0)
                            {
                                var node = pending.Pop();
                                if (!visited.Add(node)) continue;
                                foreach (var child in node.Children) pending.Push(child);
                                Interlocked.Increment(ref attemptedProviderDisconnects);
                                int result = UiaDisconnectProvider(node);
                                if (result < 0)
                                    throw new COMException("UIA disconnect failed for '" + node.Name +
                                        "', native identity " + node.NativeHandle + ", root " + (node.Parent == null) + ".", result);
                                Interlocked.Increment(ref completedProviderDisconnects);
                                DisconnectedProviderCount++;
                            }
                        }
                        catch (Exception error)
                        { teardownError = teardownError == null ? error : new AggregateException(teardownError, error); }
                        finally
                        {
                            Volatile.Write(ref teardownStage, "dispatcher disposal and owner exit");
                            Root.NativeHandle = simulatedNativeHandle;
                            Root.Window = IntPtr.Zero;
                            dispatcher?.Dispose();
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }));
                bool exitedInTime = thread.Join(TimeSpan.FromSeconds(5));
                string deadlineObservation = !exitedInTime
                    ? " Stage: " + Volatile.Read(ref teardownStage) +
                        "; attempted provider disconnects: " + Volatile.Read(ref attemptedProviderDisconnects) +
                        "; completed provider disconnects: " + Volatile.Read(ref completedProviderDisconnects) +
                        "; observed milliseconds: " + teardownClock.ElapsedMilliseconds + "."
                    : "";
                // Preserve the failed deadline, but retain the same thread until terminal cleanup.
                // A next fixture must not start alongside an unobserved owner from this one.
                if (!exitedInTime) thread.Join();
                ready.Dispose();
                if (!exitedInTime) throw new TimeoutException("Owned UIA provider did not exit within its original deadline." + deadlineObservation, teardownError);
                if (teardownError != null) throw new InvalidOperationException("Owned UIA provider teardown failed.", teardownError);
            }
            /// <summary>Exécute une action sur le thread du vrai dialogue détenu par la fixture.</summary>
            public void Invoke(Action<IntPtr> action)
            {
                (dispatcher ?? (Control)form).Invoke(new Action(() =>
                {
                    action(Handle);
                    // A native style change can recreate the WinForms HWND on this owner thread.
                    if (form != null) { Handle = form.Handle; Root.Window = Handle; }
                }));
            }
            /// <summary>Observe les notifications du vrai HWND parent détenu par ce provider.</summary>
            public Action<Message> NativeMessage { set { if (dialog != null) dialog.NativeMessage = value; else form.NativeMessage = value; } }
        }

        /// <summary>Subclass UIA d'un dialogue créé directement par Win32; sa classe reste réellement #32770.</summary>
        private sealed class ProviderDialog : NativeWindow
        {
            private readonly AutomationNode provider;
            internal bool Disconnecting;
            internal Action<Message> NativeMessage;
            internal ProviderDialog(AutomationNode node, IntPtr window) { provider = node; AssignHandle(window); }
            internal void Release() { if (Handle != IntPtr.Zero) ReleaseHandle(); }
            protected override void WndProc(ref Message message)
            {
                NativeMessage?.Invoke(message);
                if (!Disconnecting && message.Msg == 0x003D && NativeProtocolInt32(message.LParam) == AutomationInteropProvider.RootObjectId)
                {
                    provider.Window = Handle;
                    message.Result = AutomationInteropProvider.ReturnRawElementProvider(Handle, message.WParam, message.LParam, provider);
                    return;
                }
                base.WndProc(ref message);
            }
        }

        private sealed class ProviderForm : Form
        {
            public AutomationNode Provider;
            internal bool Disconnecting;
            internal bool NoActivate;
            protected override bool ShowWithoutActivation => NoActivate;
            /// <summary>Observe uniquement les notifications du parent natif détenu pour simuler les effets documentés de son propriétaire.</summary>
            public Action<Message> NativeMessage;
            protected override void WndProc(ref Message message)
            {
                NativeMessage?.Invoke(message);
                if (!Disconnecting && message.Msg == 0x003D && NativeProtocolInt32(message.LParam) == AutomationInteropProvider.RootObjectId)
                {
                    Provider.Window = Handle;
                    message.Result = AutomationInteropProvider.ReturnRawElementProvider(Handle, message.WParam, message.LParam, Provider);
                    return;
                }
                base.WndProc(ref message);
            }
        }
    }
}
