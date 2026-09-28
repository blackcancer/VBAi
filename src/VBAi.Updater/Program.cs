using System;
using System.IO;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--check-webview2")
        { try { Environment.ExitCode = WebViewRuntimePrerequisite.Installed() ? 0 : 1; } catch { Environment.ExitCode = 2; } return; }
        if (args.Length == 1 && args[0] == "--ensure-webview2")
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            using (var window = new UpdateProgressWindow()) { window.ConfigureWebView(); Application.Run(window); }
            return;
        }
        bool background = args.Length == 1 && args[0] == "--background";
        if (args.Length > 1 || (args.Length == 1 && !background)) return;
        using (var mutex = new Mutex(true, "Local\\VBAi.UpdateWorker." + WindowsIdentity.GetCurrent().User.Value, out bool owner))
        {
            if (!owner) return;
            try
            {
                var job = UpdateInstallJob.Load(UpdatePaths.Root);
                if (job == null || job.Completed) return;
                if (job.Status == "Installing update…")
                {
                    job.Completed = true; job.Status = "Installation status is uncertain. Check the installed version."; job.Save(UpdatePaths.Root);
                    if (background) return;
                }
                Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
                using (var window = new UpdateProgressWindow())
                {
                    window.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    window.Configure(UpdatePaths.Root, job, background); Application.Run(window);
                }
            }
            catch (Exception)
            {
                try { UpdatePaths.WriteAtomic(Path.Combine(UpdatePaths.Root, "worker-error.txt"), "The updater could not start."); } catch (Exception) { }
                if (!background) MessageBox.Show(UpdateText.Get("Unable to schedule the update."), "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { mutex.ReleaseMutex(); }
        }
    }
}
