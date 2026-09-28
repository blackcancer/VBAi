using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

// Diagnostic only. No input, focus change, redraw request, or full-screen capture.
public static class VbeToolbarPaintProbe
{
    [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; public Point(int x, int y) { X = x; Y = y; } }
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rectangle);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);

    private sealed class Run
    {
        public IntPtr Window;
        public uint Process;
        public string Directory, WindowClass, StartedUtc;
        public Thread Worker;
        public readonly ManualResetEvent Stop = new ManualResetEvent(false);
        public readonly Stopwatch Clock = new Stopwatch();
        public Bitmap MinimumImage, MaximumImage;
        public double Minimum = double.PositiveInfinity, Maximum = double.NegativeInfinity;
        public long MinimumAt, MaximumAt, Duration;
        public Rect MinimumBounds, MaximumBounds, PreviousBounds;
        public bool HasBounds;
        public int Attempts, Samples, ForegroundSamples, BackgroundSamples, ActivationChangedSamples, OcclusionSkipped, VisibilitySkipped, GeometrySkipped, GeometryChanges, CaptureErrors;
        public bool MinimumForeground, MaximumForeground;
        public readonly List<string> Errors = new List<string>();
        public string FirstOcclusion;
    }

    private static readonly object Gate = new object();
    private static Run active;

    public static void Begin(IntPtr window, string outputDirectory)
    {
        lock (Gate)
        {
            if (active != null) throw new InvalidOperationException("A toolbar capture is already active. Call End first.");
            if (window == IntPtr.Zero || !IsWindow(window)) throw new ArgumentException("A live toolbar HWND is required.", "window");
            if (String.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentException("An artifact directory is required.", "outputDirectory");
            var className = new StringBuilder(128);
            GetClassName(window, className, className.Capacity);
            string name = className.ToString();
            if (name != "MsoCommandBar" && name != "MsoCommandBarDock" && name != "ToolbarWindow32")
                throw new ArgumentException("Capture is restricted to a native toolbar or toolbar dock.", "window");
            uint process;
            GetWindowThreadProcessId(window, out process);
            var run = new Run
            {
                Window = window,
                Process = process,
                Directory = Path.GetFullPath(outputDirectory),
                WindowClass = name,
                StartedUtc = DateTime.UtcNow.ToString("o")
            };
            run.Worker = new Thread(delegate() { Capture(run); });
            run.Worker.IsBackground = true;
            run.Worker.Name = "VBE toolbar paint capture";
            active = run;
            run.Worker.Start();
        }
    }

    // A timeout leaves the same run available for another End call. It never
    // aborts a screenshot thread, races its Bitmap ownership, or restarts it.
    public static string End()
    {
        lock (Gate)
        {
            if (active == null) throw new InvalidOperationException("No toolbar capture is active.");
            Run run = active;
            run.Stop.Set();
            if (!run.Worker.Join(2000))
                return new JavaScriptSerializer().Serialize(new Dictionary<string, object>
                {
                    { "Completed", false }, { "Error", "The capture worker is still finishing. Call End again for this same run." }
                });
            string minimumPath = null, maximumPath = null;
            try
            {
                if (run.Samples != 0)
                {
                    Directory.CreateDirectory(run.Directory);
                    minimumPath = Path.Combine(run.Directory, "toolbar-min-light.png");
                    maximumPath = Path.Combine(run.Directory, "toolbar-max-light.png");
                    // Refuse to overwrite an earlier diagnostic's evidence.
                    if (File.Exists(minimumPath) || File.Exists(maximumPath)) throw new IOException("Toolbar capture artifacts already exist in the output directory.");
                    run.MinimumImage.Save(minimumPath, ImageFormat.Png);
                    run.MaximumImage.Save(maximumPath, ImageFormat.Png);
                }
                return new JavaScriptSerializer().Serialize(new Dictionary<string, object>
                {
                    { "Completed", true }, { "Window", "0x" + run.Window.ToInt64().ToString("X") },
                    { "ProcessId", run.Process }, { "WindowClass", run.WindowClass }, { "StartedUtc", run.StartedUtc },
                    { "MaximumDurationMs", 5000 }, { "RequestedIntervalMs", 18 }, { "DurationMs", run.Duration },
                    { "Attempts", run.Attempts }, { "ValidSamples", run.Samples },
                    { "ForegroundRequired", false }, { "ForegroundSamples", run.ForegroundSamples }, { "BackgroundSamples", run.BackgroundSamples },
                    { "ActivationChangedSamples", run.ActivationChangedSamples }, { "OcclusionSkipped", run.OcclusionSkipped },
                    { "VisibilitySkipped", run.VisibilitySkipped }, { "GeometrySkipped", run.GeometrySkipped },
                    { "GeometryChanges", run.GeometryChanges }, { "CaptureErrors", run.CaptureErrors }, { "Errors", run.Errors.ToArray() },
                    { "FirstOcclusion", run.FirstOcclusion },
                    { "LightPixelRule", "R > 170 and G > 170 and B > 170" },
                    { "MinimumLightFraction", run.Samples == 0 ? (object)null : run.Minimum },
                    { "MaximumLightFraction", run.Samples == 0 ? (object)null : run.Maximum },
                    { "MinimumAtElapsedMs", run.Samples == 0 ? (object)null : run.MinimumAt },
                    { "MaximumAtElapsedMs", run.Samples == 0 ? (object)null : run.MaximumAt },
                    { "MinimumWasForeground", run.Samples == 0 ? (object)null : run.MinimumForeground },
                    { "MaximumWasForeground", run.Samples == 0 ? (object)null : run.MaximumForeground },
                    { "MinimumBounds", run.Samples == 0 ? null : Bounds(run.MinimumBounds) },
                    { "MaximumBounds", run.Samples == 0 ? null : Bounds(run.MaximumBounds) },
                    { "MinimumImage", minimumPath }, { "MaximumImage", maximumPath },
                    { "Interpretation", run.Samples == 0 ? "No unobscured sample: no conclusion about blinking."
                        : "Sampled toolbar pixels only. Extremes can reveal a flash; their absence does not prove flicker-free rendering." },
                    { "ActivationClassification", "ForegroundSamples requires the target process to own the foreground window both before and after capture; all other valid frames are BackgroundSamples. ActivationChangedSamples counts transitions during a valid capture." },
                    { "VisibilityCheck", "A 3x3 grid of WindowFromPoint checks before and after capture, with no foreground requirement; partial occlusion between points is still possible." }
                });
            }
            finally
            {
                if (run.MinimumImage != null) run.MinimumImage.Dispose();
                if (run.MaximumImage != null) run.MaximumImage.Dispose();
                run.Stop.Dispose();
                active = null;
            }
        }
    }

    private static Dictionary<string, int> Bounds(Rect rectangle)
    {
        return new Dictionary<string, int> { { "Left", rectangle.Left }, { "Top", rectangle.Top }, { "Right", rectangle.Right }, { "Bottom", rectangle.Bottom } };
    }

    private static void Capture(Run run)
    {
        run.Clock.Start();
        try
        {
            while (!run.Stop.WaitOne(0) && run.Clock.ElapsedMilliseconds < 5000)
            {
                long started = run.Clock.ElapsedMilliseconds;
                run.Attempts++;
                Bitmap bitmap = null;
                try
                {
                    Rect bounds;
                    int validity = Validate(run, out bounds);
                    if (validity != 0) { Rejected(run, validity); continue; }
                    bool foregroundBefore = IsForegroundProcess(run.Process);
                    if (run.HasBounds && !Equal(run.PreviousBounds, bounds)) run.GeometryChanges++;
                    run.PreviousBounds = bounds;
                    run.HasBounds = true;
                    bitmap = new Bitmap(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top, PixelFormat.Format24bppRgb);
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                        graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bitmap.Size, CopyPixelOperation.SourceCopy);
                    Rect after;
                    validity = Validate(run, out after);
                    if (validity != 0) { Rejected(run, validity); continue; }
                    if (!Equal(bounds, after)) { run.GeometrySkipped++; continue; }
                    bool foregroundAfter = IsForegroundProcess(run.Process);
                    bool foreground = foregroundBefore && foregroundAfter;
                    double fraction = LightFraction(bitmap);
                    run.Samples++;
                    if (foreground) run.ForegroundSamples++; else run.BackgroundSamples++;
                    if (foregroundBefore != foregroundAfter) run.ActivationChangedSamples++;
                    long elapsed = run.Clock.ElapsedMilliseconds;
                    if (fraction < run.Minimum)
                    {
                        if (run.MinimumImage != null) run.MinimumImage.Dispose();
                        run.MinimumImage = (Bitmap)bitmap.Clone();
                        run.Minimum = fraction; run.MinimumAt = elapsed; run.MinimumBounds = bounds; run.MinimumForeground = foreground;
                    }
                    if (fraction > run.Maximum)
                    {
                        if (run.MaximumImage != null) run.MaximumImage.Dispose();
                        run.MaximumImage = (Bitmap)bitmap.Clone();
                        run.Maximum = fraction; run.MaximumAt = elapsed; run.MaximumBounds = bounds; run.MaximumForeground = foreground;
                    }
                }
                catch (Exception error)
                {
                    run.CaptureErrors++;
                    if (run.Errors.Count < 5) run.Errors.Add(error.GetType().Name + ": " + error.Message);
                }
                finally
                {
                    if (bitmap != null) bitmap.Dispose();
                    int remaining = 18 - (int)(run.Clock.ElapsedMilliseconds - started);
                    if (remaining > 0) run.Stop.WaitOne(remaining);
                }
            }
        }
        finally { run.Duration = run.Clock.ElapsedMilliseconds; run.Clock.Stop(); }
    }

    private static int Validate(Run run, out Rect bounds)
    {
        bounds = new Rect();
        if (!IsWindow(run.Window) || !IsWindowVisible(run.Window) || !GetWindowRect(run.Window, out bounds)) return 1;
        int width = bounds.Right - bounds.Left, height = bounds.Bottom - bounds.Top;
        int screenX = GetSystemMetrics(76), screenY = GetSystemMetrics(77);
        if (width < 4 || height < 4 || width > 16384 || height > 512 || (long)width * height > 4194304 ||
            bounds.Left < screenX || bounds.Top < screenY || bounds.Right > screenX + GetSystemMetrics(78) || bounds.Bottom > screenY + GetSystemMetrics(79)) return 4;
        IntPtr root = GetAncestor(run.Window, 2);
        // Capture only the requested toolbar, not a popup or another pane of the
        // same host that happens to cover it. Sampling cannot certify every pixel.
        for (int row = 0; row < 3; row++)
            for (int column = 0; column < 3; column++)
            {
                int x = bounds.Left + 1 + (width - 3) * column / 2;
                int y = bounds.Top + 1 + (height - 3) * row / 2;
                IntPtr hit = WindowFromPoint(new Point(x, y));
                if ((hit != run.Window && !IsChild(run.Window, hit)) || GetAncestor(hit, 2) != root)
                {
                    if (run.FirstOcclusion == null)
                    {
                        var name = new StringBuilder(128);
                        GetClassName(hit, name, name.Capacity);
                        uint hitProcess;
                        GetWindowThreadProcessId(hit, out hitProcess);
                        run.FirstOcclusion = "Point=" + x + "," + y + "; HWND=0x" + hit.ToInt64().ToString("X") +
                            "; class=" + name + "; process=" + hitProcess;
                    }
                    return 3;
                }
            }
        return 0;
    }

    private static void Rejected(Run run, int reason)
    {
        if (reason == 1) run.VisibilitySkipped++;
        else if (reason == 3) run.OcclusionSkipped++;
        else run.GeometrySkipped++;
    }

    private static bool IsForegroundProcess(uint targetProcess)
    {
        uint process;
        GetWindowThreadProcessId(GetForegroundWindow(), out process);
        return process == targetProcess;
    }

    private static bool Equal(Rect first, Rect second)
    {
        return first.Left == second.Left && first.Top == second.Top && first.Right == second.Right && first.Bottom == second.Bottom;
    }

    private static double LightFraction(Bitmap bitmap)
    {
        var rectangle = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        BitmapData pixels = bitmap.LockBits(rectangle, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        long light = 0;
        try
        {
            var row = new byte[bitmap.Width * 3];
            for (int y = 0; y < bitmap.Height; y++)
            {
                Marshal.Copy(IntPtr.Add(pixels.Scan0, y * pixels.Stride), row, 0, row.Length);
                for (int x = 0; x < row.Length; x += 3)
                    if (row[x] > 170 && row[x + 1] > 170 && row[x + 2] > 170) light++;
            }
        }
        finally { bitmap.UnlockBits(pixels); }
        return light / (double)((long)bitmap.Width * bitmap.Height);
    }
}
