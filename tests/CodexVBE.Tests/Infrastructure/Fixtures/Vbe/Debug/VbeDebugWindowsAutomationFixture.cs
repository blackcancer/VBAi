namespace CodexVBE.Tests.Unit
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
            public int FocusCount;
            public int SelectionCount;
            public ToggleState ToggleState;
            public double Number = 25;
            public IntPtr Window;
            public Action SelectedAction;

            public AutomationNode Add(AutomationNode child) { child.Parent = this; Children.Add(child); return child; }
            public AutomationNode With(params AutomationPattern[] patterns) { foreach (var pattern in patterns) Patterns.Add(pattern.Id); return this; }
            private AutomationNode Root { get { var root = this; while (root.Parent != null) root = root.Parent; return root; } }
            public ProviderOptions ProviderOptions => ProviderOptions.ServerSideProvider;
            public IRawElementProviderSimple HostRawElementProvider => Parent == null ? AutomationInteropProvider.HostProviderFromHandle(Window) : null;
            public object GetPatternProvider(int patternId) { return Patterns.Contains(patternId) ? this : null; }
            public object GetPropertyValue(int propertyId)
            {
                if (propertyId == AutomationElementIdentifiers.NameProperty.Id) { if (FailName) throw new ElementNotAvailableException(); return Name; }
                if (propertyId == AutomationElementIdentifiers.ControlTypeProperty.Id) return Kind.Id;
                if (propertyId == AutomationElementIdentifiers.IsEnabledProperty.Id) return Enabled;
                if (propertyId == AutomationElementIdentifiers.IsOffscreenProperty.Id) return Offscreen;
                if (propertyId == AutomationElementIdentifiers.IsPasswordProperty.Id) return Password;
                if (propertyId == AutomationElementIdentifiers.IsControlElementProperty.Id || propertyId == AutomationElementIdentifiers.IsContentElementProperty.Id) return true;
                if (propertyId == AutomationElementIdentifiers.IsKeyboardFocusableProperty.Id) return true;
                if (propertyId == AutomationElementIdentifiers.NativeWindowHandleProperty.Id) return Parent == null ? Window.ToInt32() : 0;
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
            public int[] GetRuntimeId() { return new[] { AutomationInteropProvider.AppendRuntimeId, id }; }
            public Rect BoundingRectangle => new Rect(0, 0, 100, 100);
            public IRawElementProviderSimple[] GetEmbeddedFragmentRoots() { return null; }
            public void SetFocus() { FocusCount++; }
            public IRawElementProviderFragmentRoot FragmentRoot => Root;
            public IRawElementProviderFragment ElementProviderFromPoint(double x, double y) { return this; }
            public IRawElementProviderFragment GetFocus() { return this; }
            bool ISelectionItemProvider.IsSelected => Selected;
            IRawElementProviderSimple ISelectionItemProvider.SelectionContainer => Parent;
            public void Select() { Selected = true; SelectionCount++; SelectedAction?.Invoke(); }
            public void AddToSelection() { Select(); }
            public void RemoveFromSelection() { Selected = false; }
            public IRawElementProviderSimple[] GetSelection() { return Children.Where(child => child.Selected).Cast<IRawElementProviderSimple>().ToArray(); }
            public bool CanSelectMultiple => false;
            public bool IsSelectionRequired => false;
            string IValueProvider.Value { get { if (FailValue) throw new InvalidOperationException("value unreadable"); return Text; } }
            bool IValueProvider.IsReadOnly => false;
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

        private sealed class AutomationHost : IDisposable
        {
            private readonly Thread thread;
            private readonly ManualResetEvent ready = new ManualResetEvent(false);
            private ProviderForm form;
            private Exception startupError;
            public readonly AutomationNode Root;
            public IntPtr Handle;
            public AutomationHost(AutomationNode root)
            {
                Root = root;
                thread = new Thread(() => {
                    try {
                        form = new ProviderForm { Provider = root, Text = "CodexVBE isolated UIA fixture", ShowInTaskbar = false };
                        form.Shown += (sender, args) => { Handle = form.Handle; root.Window = Handle; ready.Set(); };
                        System.Windows.Forms.Application.Run(form);
                    }
                    catch (Exception error) { startupError = error; ready.Set(); }
                }) { IsBackground = true };
                thread.SetApartmentState(ApartmentState.STA); thread.Start();
                if (!ready.WaitOne(TimeSpan.FromSeconds(10))) throw new TimeoutException("Isolated UIA provider did not start.");
                if (startupError != null) throw new InvalidOperationException("Isolated UIA provider failed.", startupError);
            }
            public void Dispose() { form.BeginInvoke(new Action(form.Close)); thread.Join(TimeSpan.FromSeconds(5)); ready.Dispose(); }
        }

        private sealed class ProviderForm : Form
        {
            public AutomationNode Provider;
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x003D && message.LParam.ToInt32() == AutomationInteropProvider.RootObjectId)
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
