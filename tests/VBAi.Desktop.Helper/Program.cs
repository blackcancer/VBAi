using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Automation;
using System.Windows.Forms;
using VBAi.Tests.Integration;

namespace VBAi.Desktop.Helper
{
    /// <summary>Runs a reviewed test script and its UI canary on an inactive desktop; never switches desktops or terminates hosts.</summary>
    internal static class Program
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SendMessageTimeoutW(IntPtr window, uint message, IntPtr wparam, IntPtr lparam, uint flags, uint milliseconds, out IntPtr result);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        private static IDisposable retainedDesktop;
        private static IsolatedTestDesktop.NativeChild retainedChild;
        private static DesktopSentinel retainedSentinel;

        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 3 && args[0] == "--probe") return Probe(args[1], args[2]);
            if (args.Length != 4 || args[0] != "--run") return 2;
            string script = Path.GetFullPath(args[1]), output = Path.GetFullPath(args[2]), worker = Path.GetFullPath(args[3]);
            if (!Path.IsPathRooted(args[1]) || !File.Exists(script) || Path.GetExtension(script) != ".ps1" ||
                !Path.IsPathRooted(args[2]) || Directory.Exists(output) ||
                !Path.IsPathRooted(args[3]) || !File.Exists(worker)) return 3;
            Directory.CreateDirectory(output);
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            string input = IsolatedTestDesktop.InputDesktopName();
            try
            {
                retainedDesktop = IsolatedTestDesktop.Create(desktop);
                retainedSentinel = new DesktopSentinel(desktop);
                retainedSentinel.Start();
                Write(output, "sentinel-started.json", new { Desktop = desktop, ProcessId = Process.GetCurrentProcess().Id,
                    ThreadId = retainedSentinel.ThreadId, Window = retainedSentinel.Window.ToInt64(),
                    Hidden = true, FocusRequested = false, ImeDisabledOnlyForSentinelThread = true, Utc = Utc() });
                Write(output, "desktop-plan.json", new { Desktop = desktop, InputDesktop = input,
                    Script = script, SwitchDesktopCalled = false, TerminationAllowed = false,
                    Scope = "Window/focus isolation under the same user/profile; not a security sandbox or native acceptance", Utc = Utc() });
                string self = typeof(Program).Assembly.Location;
                retainedChild = IsolatedTestDesktop.Launch(self,
                    new[] { "--probe", desktop, Path.Combine(output, "canary.json") }, Path.GetDirectoryName(script), desktop);
                Observe(output, desktop, "canary", 30000);
                uint canaryExit = retainedChild.ExitCode();
                retainedChild.Dispose(); retainedChild = null;
                if (canaryExit != 0) throw new InvalidOperationException("Inactive-desktop UI canary failed; test script was not started.");
                string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
                retainedChild = IsolatedTestDesktop.Launch(powershell,
                    new[] { "-NoProfile", "-NonInteractive", "-File", worker, "-ScriptPath", script,
                        "-ExpectedDesktop", desktop, "-ProofPath", Path.Combine(output, "worker-desktop.json"),
                        "-HelperAssembly", self, "-SentinelWindow", retainedSentinel.Window.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture),
                        "-SentinelProcessId", Process.GetCurrentProcess().Id.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        "-SentinelThreadId", retainedSentinel.ThreadId.ToString(System.Globalization.CultureInfo.InvariantCulture) },
                    Path.GetDirectoryName(script), desktop);
                Observe(output, desktop, "campaign", 1800000);
                uint code = retainedChild.ExitCode();
                if (IsolatedTestDesktop.HasOtherWindows(desktop, retainedSentinel.Window, retainedSentinel.ThreadId,
                    (uint)Process.GetCurrentProcess().Id))
                    throw new InvalidOperationException("The script exited while private desktop windows remain; ownership is retained without cleanup replay.");
                Write(output, "sentinel-inventory.json", new { Desktop = desktop, SentinelObserved = true,
                    SentinelWindow = retainedSentinel.Window.ToInt64(), OtherWindows = 0,
                    OriginalChildExitObserved = true, Utc = Utc() });
                retainedSentinel.CloseOnce();
                Write(output, "sentinel-exit.json", new { Desktop = desktop, OriginalThreadExitObserved = true,
                    NativeCloseAttempts = 1, DesktopSwitches = 0, Utc = Utc() });
                retainedSentinel = null;
                retainedChild.Dispose(); retainedChild = null;
                retainedDesktop.Dispose(); retainedDesktop = null;
                Write(output, "terminal.json", new { State = "ORIGINAL_CHILD_EXIT_OBSERVED", ExitCode = code,
                    Desktop = desktop, InputDesktop = IsolatedTestDesktop.InputDesktopName(),
                    DesktopSwitches = 0, OwnedForegroundObservations = 0, Utc = Utc() });
                return unchecked((int)code);
            }
            catch (Exception error)
            {
                Write(output, "failure.json", new { State = retainedChild == null && retainedSentinel == null ? "REFUSED" : "RETAINED_UNCERTAIN",
                    Desktop = desktop, ProcessId = retainedChild == null ? 0 : retainedChild.ProcessId,
                    SentinelWindow = retainedSentinel == null ? 0L : retainedSentinel.Window.ToInt64(),
                    Error = error.ToString(), CleanupReplayed = false, Utc = Utc() });
                // A UI/exit uncertainty is not permission to stop a host or drop its original query handle.
                if (retainedChild != null || retainedSentinel != null)
                    for (;;) Thread.Sleep(1000);
                if (retainedDesktop != null) { retainedDesktop.Dispose(); retainedDesktop = null; }
                return 1;
            }
        }

        /// <summary>One hidden raw HWND makes final desktop enumeration positively verifiable without focus or input injection.</summary>
        private sealed class DesktopSentinel
        {
            private const uint CloseMessage = 0x8000 + 73;
            [DllImport("imm32.dll")] private static extern bool ImmDisableIME(uint threadId);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr CreateWindowExW(uint extendedStyle, string className, string title, uint style,
                int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
            [DllImport("user32.dll", SetLastError = true)] private static extern bool DestroyWindow(IntPtr window);
            [DllImport("user32.dll", SetLastError = true)] private static extern int GetMessageW(out NativeMessage message, IntPtr window, uint minimum, uint maximum);
            [DllImport("user32.dll")] private static extern bool TranslateMessage(ref NativeMessage message);
            [DllImport("user32.dll")] private static extern IntPtr DispatchMessageW(ref NativeMessage message);
            [DllImport("user32.dll", SetLastError = true)] private static extern bool PostThreadMessageW(uint thread, uint message, IntPtr wparam, IntPtr lparam);
            [StructLayout(LayoutKind.Sequential)]
            private struct NativeMessage
            {
                internal IntPtr Window;
                internal uint Message;
                internal UIntPtr Wparam;
                internal IntPtr Lparam;
                internal uint Time;
                internal int X, Y;
                internal uint Private;
            }
            private readonly string desktop;
            private readonly ManualResetEventSlim ready = new ManualResetEventSlim();
            private Thread thread;
            private IDisposable desktopLease;
            private Exception failure;
            private int closeRequested;
            private bool destroyed;
            internal IntPtr Window { get; private set; }
            internal uint ThreadId { get; private set; }
            internal DesktopSentinel(string desktop) { this.desktop = desktop; }

            internal void Start()
            {
                thread = new Thread(Run) { IsBackground = true };
                // This thread uses only Win32; MTA avoids an implicit CLR STA/OLE window.
                thread.SetApartmentState(ApartmentState.MTA);
                thread.Start();
                if (!ready.Wait(10000) || failure != null || Window == IntPtr.Zero)
                    throw new InvalidOperationException("The private desktop sentinel did not establish its exact native identity; no child is launched.", failure);
            }

            private void Run()
            {
                try
                {
                    desktopLease = IsolatedTestDesktop.BindPrivateSentinelThread(desktop);
                    IsolatedTestDesktop.RequireCurrent(desktop);
                    ThreadId = IsolatedTestDesktop.GetCurrentThreadId();
                    // Windows otherwise creates an extra IME HWND for this thread. This raw inventory
                    // sentinel never accepts input; disable IME only on its own fresh thread before WM_CREATE.
                    if (ThreadId == 0 || !ImmDisableIME(ThreadId))
                        throw new InvalidOperationException("The sentinel-only IME context could not be disabled before its first window; no child is launched.");
                    // STATIC is a Windows class. No ShowWindow, activation, WinForms parking window or UIA action is used.
                    Window = CreateWindowExW(0x08000000, "STATIC", "Owned VBAi private desktop inventory sentinel", // WS_EX_NOACTIVATE
                        0x80000000, 0, 0, 1, 1, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                    uint pid; uint ownerThread = GetWindowThreadProcessId(Window, out pid);
                    if (Window == IntPtr.Zero || ownerThread != ThreadId || pid != (uint)Process.GetCurrentProcess().Id)
                        throw new InvalidOperationException("The sentinel HWND ownership differs from its exact creating thread.");
                    ready.Set();
                    NativeMessage message;
                    int result;
                    while ((result = GetMessageW(out message, IntPtr.Zero, 0, 0)) > 0)
                    {
                        if (message.Window == IntPtr.Zero && message.Message == CloseMessage)
                        {
                            IsolatedTestDesktop.RequireCurrent(desktop);
                            uint currentPid; uint currentThread = GetWindowThreadProcessId(Window, out currentPid);
                            if (message.Lparam != Window || Volatile.Read(ref closeRequested) != 1 || currentThread != ThreadId ||
                                currentPid != (uint)Process.GetCurrentProcess().Id || IsolatedTestDesktop.DesktopName(currentThread) != desktop)
                                throw new InvalidOperationException("Sentinel close identity changed; its native close is refused.");
                            if (!DestroyWindow(Window))
                                throw new InvalidOperationException("The single sentinel DestroyWindow attempt failed; it is never replayed.");
                            destroyed = true;
                            break;
                        }
                        TranslateMessage(ref message); DispatchMessageW(ref message);
                    }
                    if (!destroyed) throw new InvalidOperationException("The sentinel message loop ended without its reviewed native close.");
                }
                catch (Exception error) { failure = error; }
                finally { ready.Set(); }
            }

            internal void CloseOnce()
            {
                if (Interlocked.CompareExchange(ref closeRequested, 1, 0) != 0)
                    throw new InvalidOperationException("The sentinel close request cannot be replayed.");
                if (failure != null || thread == null || !thread.IsAlive ||
                    !PostThreadMessageW(ThreadId, CloseMessage, IntPtr.Zero, Window))
                    throw new InvalidOperationException("The original sentinel thread could not receive its single close request.", failure);
                if (!thread.Join(5000) || failure != null || !destroyed)
                    throw new InvalidOperationException("The original sentinel thread did not complete its native close; ownership is retained.", failure);
                uint pid;
                if (GetWindowThreadProcessId(Window, out pid) != 0)
                    throw new InvalidOperationException("The sentinel HWND still exists after observed thread exit; desktop release is refused.");
                desktopLease.Dispose(); desktopLease = null;
                ready.Dispose();
            }
        }

        private static void Observe(string output, string desktop, string phase, int deadline)
        {
            Write(output, phase + "-started.json", new { ProcessId = retainedChild.ProcessId,
                ThreadId = retainedChild.ThreadId, OriginalHandle = retainedChild.ProcessHandle.ToInt64(), Utc = Utc() });
            var watch = Stopwatch.StartNew(); bool pending = false;
            while (!retainedChild.Wait(200))
            {
                if (string.Equals(IsolatedTestDesktop.InputDesktopName(), desktop, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The isolated desktop became the input desktop; campaign ownership is retained.");
                uint pid; GetWindowThreadProcessId(GetForegroundWindow(), out pid);
                if (pid == retainedChild.ProcessId)
                    throw new InvalidOperationException("The owned launcher appeared in the input foreground; isolation refused.");
                if (!pending && watch.ElapsedMilliseconds >= deadline)
                {
                    pending = true;
                    Write(output, phase + "-pending.json", new { State = "PENDING_RETAINED", ProcessId = retainedChild.ProcessId,
                        OriginalHandle = retainedChild.ProcessHandle.ToInt64(), ObservationBoundMs = deadline,
                        NoRelaunch = true, NoTermination = true, Utc = Utc() });
                }
            }
            Write(output, phase + "-exit.json", new { ProcessId = retainedChild.ProcessId,
                OriginalHandle = retainedChild.ProcessHandle.ToInt64(), ExitCode = retainedChild.ExitCode(),
                WaitElapsedMs = watch.ElapsedMilliseconds, PendingPreviouslyRecorded = pending, Utc = Utc() });
        }

        private static int Probe(string expected, string receipt)
        {
            try
            {
                IsolatedTestDesktop.RequireCurrent(expected);
                bool clicked = false, expired = false; Exception failure = null; Thread reader = null;
                object observation = null;
                using (var form = new Form { Text = "Disposable VBAi isolated-desktop canary", Width = 400, Height = 180 })
                using (var button = new Button { Name = "ownedCanary", Text = "Owned canary", Dock = DockStyle.Fill })
                using (var timer = new System.Windows.Forms.Timer { Interval = 10000 })
                {
                    form.Controls.Add(button);
                    timer.Tick += (sender, value) => { expired = true; timer.Stop(); form.Close(); };
                    button.Click += (sender, value) => { clicked = true; timer.Stop(); form.Close(); };
                    form.Shown += (sender, value) => {
                        button.Focus(); IntPtr hwnd = button.Handle; uint nativePid;
                        uint nativeThread = GetWindowThreadProcessId(hwnd, out nativePid);
                        IsolatedTestDesktop.RequireCurrent(expected);
                        if (IsolatedTestDesktop.DesktopName(nativeThread) != expected)
                            throw new InvalidOperationException("The canary window is on a different desktop.");
                        reader = new Thread(() => {
                            try
                            {
                                IsolatedTestDesktop.RequireCurrent(expected);
                                var element = AutomationElement.FromHandle(hwnd);
                                if (element.Current.ProcessId != Process.GetCurrentProcess().Id ||
                                    element.Current.AutomationId != "ownedCanary" || element.Current.ControlType != ControlType.Button)
                                    throw new InvalidOperationException("The private UIA canary identity differs.");
                                observation = new { Desktop = expected, InputDesktop = IsolatedTestDesktop.InputDesktopName(),
                                    NativeProcessId = nativePid, NativeThreadId = nativeThread,
                                    NativeHandle = hwnd.ToInt64(), UiAProcessId = element.Current.ProcessId,
                                    UiAOffscreen = element.Current.IsOffscreen, UiAEnabled = element.Current.IsEnabled };
                                // The stock UIA WindowsButton proxy uses SendInput internally. Private-desktop
                                // acceptance uses exact native HWND dispatch, without keyboard/focus injection.
                                uint clickPid; uint clickThread = GetWindowThreadProcessId(hwnd, out clickPid);
                                if (clickPid != nativePid || clickThread != nativeThread ||
                                    IsolatedTestDesktop.DesktopName(clickThread) != expected)
                                    throw new InvalidOperationException("The canary changed identity before its single native click.");
                                IntPtr clickResult;
                                if (SendMessageTimeoutW(hwnd, 0x00F5, IntPtr.Zero, IntPtr.Zero, 0x0002, 3000, out clickResult) == IntPtr.Zero)
                                    throw new InvalidOperationException("The single native canary click did not return; no click is replayed.");
                            }
                            catch (Exception error) { failure = error; try { form.BeginInvoke((Action)form.Close); } catch { } }
                        }) { IsBackground = true };
                        reader.SetApartmentState(ApartmentState.MTA); reader.Start(); timer.Start();
                    };
                    Application.Run(form);
                    if (reader == null || !reader.Join(3000) || expired || failure != null || !clicked)
                        throw new InvalidOperationException("Private desktop WinForms/UIA canary did not complete.", failure);
                }
                File.WriteAllText(receipt, new JavaScriptSerializer().Serialize(new { State = "PASS_CANARY_ONLY",
                    Observation = observation, ClickObserved = clicked, DesktopSwitches = 0, Utc = Utc() }));
                return 0;
            }
            catch (Exception error)
            {
                if (!File.Exists(receipt)) File.WriteAllText(receipt, new JavaScriptSerializer().Serialize(new {
                    State = "CANARY_FAILED", Error = error.ToString(), Utc = Utc() }));
                return 1;
            }
        }

        private static string Utc() => DateTime.UtcNow.ToString("o");
        private static void Write(string root, string name, object value)
        {
            using (var file = new FileStream(Path.Combine(root, name), FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(file, new System.Text.UTF8Encoding(false)))
                writer.Write(new JavaScriptSerializer().Serialize(value));
        }
    }
}
