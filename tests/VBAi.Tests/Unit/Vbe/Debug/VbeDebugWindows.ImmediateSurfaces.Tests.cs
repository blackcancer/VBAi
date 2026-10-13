using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Windows.Automation;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        [DataRow("Document", false), DataRow("Document", true)]
        [DataRow("Edit", false), DataRow("Edit", true)]
        [DataRow("Custom", false), DataRow("Custom", true)]
        [DataRow("Pane", false), DataRow("Pane", true)]
        public void ImmediateTextSurfaceUsesItsCapabilitiesAtTheRootOrInsideThePane(string kind, bool child)
        {
            ControlType type = kind == "Document" ? ControlType.Document : kind == "Edit" ? ControlType.Edit :
                kind == "Custom" ? ControlType.Custom : ControlType.Pane;
            var text = new AutomationNode { Kind = type, Name = "Owned Immediate", Text = "previous output\r\n" }
                .With(TextPattern.Pattern);
            var root = child ? new AutomationNode { Kind = ControlType.Pane, Name = "Owned pane" } : text;
            if (child) root.Add(text);
            using (var host = new AutomationHost(root, noActivate: true))
            {
                int nativeFocus = 0;
                host.NativeMessage = message =>
                {
                    if (message.Msg == 0x0007) Interlocked.Increment(ref nativeFocus); // WM_SETFOCUS on the owned HWND.
                };
                var native = Native<VbeDebugWindows.IImmediateProbe>("NativeImmediateProbe");
                Assert.AreEqual(text.Text, native.Prepare(host.Handle));
                Assert.AreEqual(text.Text, native.Text(host.Handle));
                // UIA uses the native HWND for roots; fragments receive the provider callback.
                Assert.AreEqual(1, child ? text.FocusCount : Volatile.Read(ref nativeFocus));
                Assert.AreEqual(1, text.SelectionCount);
            }
        }

        [TestMethod]
        [DataRow("documents"), DataRow("edits"), DataRow("mixed fallback"), DataRow("unsupported")]
        public void ImmediateTextSurfaceRefusesAmbiguousOrUnsupportedProvidersBeforeFocus(string scenario)
        {
            var root = new AutomationNode { Kind = ControlType.Pane, Name = "Owned pane" };
            var first = root.Add(new AutomationNode { Kind = scenario == "documents" || scenario == "unsupported" ?
                ControlType.Document : ControlType.Edit, Name = "First" });
            if (scenario != "unsupported") first.With(TextPattern.Pattern);
            if (scenario != "unsupported") root.Add(new AutomationNode
            {
                Kind = scenario == "documents" ? ControlType.Document : scenario == "edits" ? ControlType.Edit : ControlType.Custom,
                Name = "Second"
            }.With(TextPattern.Pattern));
            using (var host = new AutomationHost(root, noActivate: true))
            {
                var native = Native<VbeDebugWindows.IImmediateProbe>("NativeImmediateProbe");
                Assert.ThrowsException<InvalidOperationException>(() => native.Prepare(host.Handle));
                Assert.AreEqual(0, root.FocusCount);
                Assert.AreEqual(0, first.FocusCount);
                Assert.AreEqual(0, first.SelectionCount);
            }
        }
    }
}
