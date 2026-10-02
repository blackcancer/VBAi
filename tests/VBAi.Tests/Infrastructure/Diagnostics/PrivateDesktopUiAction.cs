using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using Accessibility;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Integration
{
    /// <summary>Delivers one action to an exact owned control on the inactive qualification desktop.</summary>
    internal static class PrivateDesktopUiAction
    {
        private const uint BmClick = 0x00F5, ObjidClient = 0xFFFFFFFC, SmtoAbortIfHung = 2;
        private static readonly Guid IAccessibleId = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr handle, StringBuilder name, int size);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr handle, uint message, IntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("oleacc.dll", EntryPoint = "AccessibleObjectFromWindow")]
        private static extern int AccessibleObjectFromWindow(IntPtr handle, uint objectId, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);

        internal static bool Enabled => !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP"));

        private static void RequireDesktop(uint nativeThread)
        {
            string required = Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP");
            string configured = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            if (string.IsNullOrEmpty(required) || !string.Equals(required, configured, StringComparison.Ordinal))
                throw new InvalidOperationException("Private UI action requires the exact worker and test desktop pair.");
            IsolatedTestDesktop.RequireCurrent(required);
            if (!string.Equals(IsolatedTestDesktop.DesktopName(nativeThread), required, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The native UI owner thread is outside the exact inactive desktop.");
        }

        internal static bool MatchesButton(string observedId, string expectedId, string observedName, string expectedName,
            string observedType, int uiProcessId, int nativeProcessId, uint nativeThreadId,
            int expectedProcessId, uint expectedThreadId, long observedHandle, long targetHandle,
            bool enabled, bool offscreen, bool nativeVisible, bool nativeEnabled, bool childOfOwner)
            => !string.IsNullOrEmpty(expectedId) && !string.IsNullOrEmpty(expectedName) &&
               string.Equals(observedId, expectedId, StringComparison.Ordinal) &&
               string.Equals(observedName, expectedName, StringComparison.Ordinal) &&
               string.Equals(observedType, ControlType.Button.ProgrammaticName, StringComparison.Ordinal) &&
               uiProcessId == expectedProcessId && nativeProcessId == expectedProcessId &&
               nativeThreadId == expectedThreadId && expectedThreadId != 0 &&
               observedHandle != 0 && observedHandle == targetHandle &&
               enabled && !offscreen && nativeVisible && nativeEnabled && childOfOwner;

        internal static bool MatchesVirtualGit(string observedName, string expectedName, string observedType,
            int uiProcessId, int popupProcessId, uint popupThreadId, int expectedProcessId, uint expectedThreadId,
            long itemHandle, long itemAncestor, long popupHandle, bool enabled, bool offscreen, bool popupVisible)
            => !string.IsNullOrEmpty(expectedName) &&
               string.Equals(observedName, expectedName, StringComparison.Ordinal) &&
               string.Equals(observedType, ControlType.MenuItem.ProgrammaticName, StringComparison.Ordinal) &&
               uiProcessId == expectedProcessId && popupProcessId == expectedProcessId &&
               popupThreadId == expectedThreadId && expectedThreadId != 0 &&
               itemHandle == 0 && popupHandle != 0 && itemAncestor == popupHandle &&
               enabled && !offscreen && popupVisible;

        internal static void ClickButtonOnce(AutomationElement item, string expectedId, string expectedName,
            IntPtr owner, IntPtr target, int expectedProcessId, uint expectedThreadId)
        {
            uint pid; uint thread = GetWindowThreadProcessId(target, out pid);
            RequireDesktop(thread);
            if (!IsWindow(target) || !IsWindow(owner) ||
                !MatchesButton(item.Current.AutomationId, expectedId, item.Current.Name, expectedName,
                    item.Current.ControlType.ProgrammaticName, item.Current.ProcessId, (int)pid, thread,
                    expectedProcessId, expectedThreadId, unchecked((long)(uint)item.Current.NativeWindowHandle),
                    unchecked((long)(uint)target.ToInt64()), item.Current.IsEnabled, item.Current.IsOffscreen,
                    IsWindowVisible(target), IsWindowEnabled(target), IsChild(owner, target)))
                throw new InvalidOperationException("The exact owned native button changed before delivery.");
            var cls = new StringBuilder(128);
            if (GetClassName(target, cls, cls.Capacity) == 0 ||
                !cls.ToString().Contains(".BUTTON."))
                throw new InvalidOperationException("The exact owned action is not a WinForms native Button.");
            IntPtr ignored;
            // SendMessageTimeout returns a delivery status; BM_CLICK has no result value.
            // A timeout/error leaves the click outcome uncertain and must never trigger a replay.
            if (SendMessageTimeout(target, BmClick, IntPtr.Zero, IntPtr.Zero, SmtoAbortIfHung, 5000, out ignored) == IntPtr.Zero)
                throw new InvalidOperationException("The exact native BM_CLICK outcome is uncertain; no replay.");
        }

        private static IntPtr NativeAncestor(AutomationElement element)
        {
            for (int depth = 0; element != null && depth < 16; depth++, element = TreeWalker.RawViewWalker.GetParent(element))
                if (element.Current.NativeWindowHandle != 0)
                    return new IntPtr(unchecked((long)(uint)element.Current.NativeWindowHandle));
            return IntPtr.Zero;
        }

        internal static void InvokeVirtualGitOnce(AutomationElement item, IntPtr popup, string expectedLabel,
            int expectedProcessId, uint expectedThreadId)
        {
            uint pid; uint thread = GetWindowThreadProcessId(popup, out pid);
            RequireDesktop(thread);
            if (!IsWindow(popup) ||
                !MatchesVirtualGit(item.Current.Name, expectedLabel, item.Current.ControlType.ProgrammaticName,
                    item.Current.ProcessId, (int)pid, thread, expectedProcessId, expectedThreadId,
                    item.Current.NativeWindowHandle, unchecked((long)(uint)NativeAncestor(item).ToInt64()),
                    unchecked((long)(uint)popup.ToInt64()),
                    item.Current.IsEnabled, item.Current.IsOffscreen, IsWindowVisible(popup)))
                throw new InvalidOperationException("The exact owned virtual Git item changed before delivery.");
            var cls = new StringBuilder(128);
            if (GetClassName(popup, cls, cls.Capacity) == 0 ||
                !NativeToolStripPopupIdentity.Matches(AutomationElement.FromHandle(popup).Current.ControlType.ProgrammaticName, cls.ToString()))
                throw new InvalidOperationException("The exact owned Git popup is not a WinForms menu.");
            Guid iid = IAccessibleId; IAccessible accessible = null;
            Exception actionError = null;
            try
            {
                int hr = AccessibleObjectFromWindow(popup, ObjidClient, ref iid, out accessible);
                if (hr != 0 || accessible == null)
                    throw new InvalidOperationException("Exact popup MSAA client unavailable: HRESULT 0x" +
                        unchecked((uint)hr).ToString("X8"));
                int count = accessible.accChildCount, exactChild = 0;
                if (count < 1 || count > 64) throw new InvalidOperationException("Bounded Git popup MSAA child inventory refused.");
                for (int child = 1; child <= count; child++)
                    if (string.Equals(accessible.get_accName(child), expectedLabel, StringComparison.Ordinal) &&
                        Convert.ToInt32(accessible.get_accRole(child)) == 12)
                    {
                        if (exactChild != 0) throw new InvalidOperationException("Ambiguous exact MSAA Git child.");
                        exactChild = child;
                    }
                if (exactChild == 0) throw new InvalidOperationException("Exact Git item absent from the owned MSAA popup.");
                // The child id is unique in this exact popup; the virtual UIA item has HWND zero.
                accessible.accDoDefaultAction(exactChild);
            }
            catch (Exception error) { actionError = error; throw; }
            finally
            {
                if (accessible != null && Marshal.IsComObject(accessible))
                    try { Marshal.ReleaseComObject(accessible); }
                    catch (Exception cleanup)
                    {
                        if (actionError != null)
                            throw new AggregateException("Virtual Git action and MSAA release errors are preserved.", actionError, cleanup);
                        throw;
                    }
            }
        }
    }
}
