namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        private static extern uint OwnedOptionsFixtureProcess(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")]
        private static extern int OwnedOptionsFixtureClass(IntPtr window, StringBuilder text, int capacity);

        /// <summary>Raccorde la scène Win32 à l'identité réelle du dialogue détenu, sans fabriquer une classe/PID pour la production.</summary>
        private static void BindOwnedOptionsDialog(SystemScene scene, AutomationHost host)
        {
            OwnedOptionsFixtureProcess(host.Handle, out uint processId);
            var text = new StringBuilder(128); OwnedOptionsFixtureClass(host.Handle, text, text.Capacity);
            Assert.AreNotEqual(IntPtr.Zero, host.Handle);
            Assert.AreEqual((uint)Process.GetCurrentProcess().Id, processId);
            Assert.AreEqual("#32770", text.ToString());
            scene.Windows.Add(new SystemWindow { Handle = host.Handle, ProcessId = processId, Class = text.ToString(), Text = host.Root.Name });
        }
    }
}
