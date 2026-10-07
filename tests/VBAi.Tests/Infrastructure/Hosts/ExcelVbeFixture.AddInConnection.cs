using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        // Only the Q028 bank opts in. /automation prevents automatic add-in loading;
        // connecting the registered exact candidate is a separate, single native action.
        private void ConnectQ028OwnedAddIn(string desktop)
        {
            if (Environment.GetEnvironmentVariable("VBAi_Q028_CONNECT_OWNED_ADDIN") != "1") return;
            bool mainDesktop = OllamaOfficeDesktop.MainEnabled;
            if (!owned || application == null || (!mainDesktop && (privateDesktopChild == null || string.IsNullOrEmpty(desktop))))
                throw new InvalidOperationException("An original owned Excel generation and explicit desktop scope are required.");
            int owningThread = Thread.CurrentThread.ManagedThreadId;
            var process = ownedProcess; var child = privateDesktopChild; var app = application;
            object editor = null, main = null, addins = null, entry = null;
            var rows = new List<IDictionary<string, object>>();
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            try
            {
                editor = ((dynamic)app).VBE; main = ((dynamic)editor).MainWindow;
                IntPtr hwnd = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; uint tid = GetWindowThreadProcessId(hwnd, out pid);
                if (pid != (uint)ProcessId || tid == 0) throw new InvalidOperationException("Original VBE window owner is absent.");
                addins = ((dynamic)editor).AddIns; entry = ((dynamic)addins).Item("VBAi.AddIn");
                Action guard = () =>
                {
                    if (Thread.CurrentThread.ManagedThreadId != owningThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                        !ReferenceEquals(app, application) || !ReferenceEquals(process, ownedProcess) || !ReferenceEquals(child, privateDesktopChild) ||
                        process.HasExited || process.Id != ProcessId || (child != null && child.ProcessId != ProcessId))
                        throw new InvalidOperationException("The original Excel/STA generation changed.");
                    OllamaOfficeDesktop.Require(desktop);
                    uint currentPid; IntPtr current = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                    uint currentTid = GetWindowThreadProcessId(current, out currentPid);
                    if (current != hwnd || currentPid != (uint)ProcessId || currentTid != tid || !IsWindowEnabled(current) ||
                        !string.Equals((string)((dynamic)entry).ProgId, "VBAi.AddIn", StringComparison.Ordinal) ||
                        new Guid((string)((dynamic)entry).Guid) != new Guid("8E854243-087F-4D6C-9E0E-8622B0E50883"))
                        throw new InvalidOperationException("The original Excel VBE/AddIn identity is not exact.");
                    OllamaOfficeDesktop.RequireWindow(desktop, (uint)ProcessId, true, current);
                };
                gate.Run("VBAi.AddIn", guard, () => ((dynamic)entry).Connect,
                    () => ((dynamic)entry).Connect = true, ReadQ028AddInRegistration, row =>
                    {
                        row["Host"] = "Excel"; row["ProcessId"] = ProcessId; row["Desktop"] = desktop;
                        row["SharedConnectionGateOrigin"] = "Qualified private Access single-use gate; applied here only to owned Excel";
                        row["OriginalProcessHandle"] = child != null ? child.ProcessHandle.ToInt64() : process.Handle.ToInt64();
                        row["ExpectedMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
                        rows.Add(row); WriteEvidence("private-excel-addin-connection.json", rows);
                    });
            }
            catch { PreserveForDiagnosticRecovery = true; throw; }
            finally { Release(entry); Release(addins); Release(main); Release(editor); }
        }

        private static string ReadQ028AddInRegistration()
        {
            var rows = new List<string>();
            foreach (RegistryHive hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
                using (var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64))
                using (var key = root.OpenSubKey(@"Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn", false))
                    rows.Add(hive + "|" + Convert.ToString(key?.GetValue("LoadBehavior")));
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            using (var server = root.OpenSubKey(@"Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32", false))
            {
                if (server == null) throw new InvalidOperationException("Exact registered VBAi server is absent.");
                var keys = new List<string> { "" }; keys.AddRange(server.GetSubKeyNames());
                foreach (string name in keys)
                {
                    Version version;
                    if (name.Length != 0 && !Version.TryParse(name, out version)) continue;
                    using (var selected = name.Length == 0 ? root.OpenSubKey(server.Name.Substring(server.Name.IndexOf('\\') + 1), false) : server.OpenSubKey(name, false))
                    {
                        string codeBase = Convert.ToString(selected?.GetValue("CodeBase"));
                        if (string.IsNullOrWhiteSpace(codeBase)) throw new InvalidOperationException("Registered candidate path is absent.");
                        string path = new Uri(codeBase).LocalPath;
                        if (HashQ028Assembly(path) != HashQ028Assembly(typeof(VbeSession).Assembly.Location))
                            throw new InvalidOperationException("Registered product bytes differ from the frozen test candidate; no Connect.");
                        rows.Add(name + "|" + codeBase);
                    }
                }
            }
            return string.Join(";", rows);
        }

        private static string HashQ028Assembly(string path)
        {
            using (var sha = SHA256.Create())
            using (var stream = System.IO.File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(stream));
        }
    }
}
