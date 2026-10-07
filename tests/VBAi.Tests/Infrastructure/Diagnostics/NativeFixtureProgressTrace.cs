using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Passive bounded fixture phases; records neither COM values nor request/document content.</summary>
    internal sealed class NativeFixtureProgressTrace : IDisposable
    {
        internal const int MaximumEvents = 1024;
        internal const int MaximumFileBytes = 1024 * 1024;
        [ThreadStatic] private static NativeFixtureProgressTrace current;
        private readonly NativeFixtureProgressTrace previous;
        private readonly Action<string> writer;
        private readonly Stopwatch clock = Stopwatch.StartNew();
        private readonly string correlation = Guid.NewGuid().ToString("N");
        private int events;
        private bool disposed;
        internal bool WriterAvailable => writer != null;

        internal enum Phase
        {
            Scope, PrepareForm, FormProject, FormComponents, FormAdd, FormName, FormDesigner, FormCaption,
            FormWidth, FormHeight, FormControls, LabelAdd, LabelCaption, LabelLeft, LabelTop, LabelWidth, LabelHeight,
            ButtonAdd, ButtonCaption, ButtonLeft, ButtonTop, ButtonWidth, ButtonHeight, FormCodeModule, FormCodeInsert,
            FormSeedFont, InitialSaveAs, ReleaseCode, ReleaseButton, ReleaseLabel, ReleaseControls, ReleaseDesigner,
            ReleaseComponent, ReleaseComponents, ReleaseProject, PrepareLayout, EventsDisable, MacrosDisable,
            LayoutDesigner, LayoutWidth, LayoutHeight, LayoutControls, LayoutAdd, LayoutGeometry, LayoutTag,
            NestedLayout, FrameCaption, FrameControls, MultiPageAdd, MultiPageGeometry, MultiPagePages, FirstPage,
            PageCaption, PageControls, LeafAdd, LeafGeometry, LeafText, ReleaseLeaf, ReleasePageControls, ReleasePage,
            ReleasePages, ReleaseMultiPage, ReleaseNestedControls, ReleaseLayoutControl,
            GeometryLeft, GeometryTop, GeometryWidth, GeometryHeight, PersistLayout, RenderDesigner,
            SeedBeforeSave, PersistSave, SeedAfterSave, PersistReopen, SeedAfterReopen,
            CaptureDesigner, CaptureProject, CaptureComponents, CaptureComponent, DesignerWindow, DesignerVisible,
            DesignerFocus, CaptureVbe, CaptureMain, PumpEvents, PumpDelay, CaptureState, ActiveProject, ActiveWindow,
            DesignerHandle, MainHandle, DesignerCaption, ActiveCaption, DesignerType, ActiveType, MainVisible,
            ProjectIdentity, DesignerIdentity, ReleaseActiveWindow, ReleaseActiveProject, SelectCaptureTarget,
            CaptureRectangle, CreateBitmap, CreateGraphics, AcquireHdc, PrintWindow, ReleaseHdc, SavePng,
            CaptureStateAfter, SelectCaptureTargetAfter, CaptureEvidence, ReleaseMain, ReleaseVbe, ReleaseWindow, LayoutProject, LayoutComponents, LayoutComponent, LayoutAction
        }
        private NativeFixtureProgressTrace(Action<string> writer)
        { this.writer = writer; previous = current; current = this; Record(Phase.Scope, "Entered", null); }

        internal static NativeFixtureProgressTrace BeginAtPath(string path)
        {
            Action<string> writer = null;
            if (!string.IsNullOrWhiteSpace(path))
            {
                try
                {
                    FormFontObservation.RequireSafeLocalPath(path, false);
                    using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) file.Flush(true);
                    writer = line =>
                    {
                        FormFontObservation.RequireSafeLocalPath(path, true);
                        byte[] bytes = new UTF8Encoding(false, true).GetBytes(line + "\n");
                        using (var file = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.Read))
                        {
                            if (file.Length > MaximumFileBytes - bytes.Length) return;
                            file.Seek(0, SeekOrigin.End); file.Write(bytes, 0, bytes.Length); file.Flush(true);
                        }
                    };
                }
                catch { /* A diagnostic writer must not change native fixture behavior. */ }
            }
            return new NativeFixtureProgressTrace(writer);
        }
        internal static NativeFixtureProgressTrace Begin(Action<string> writer) => new NativeFixtureProgressTrace(writer);

        internal static T Read<T>(Phase phase, Func<T> action)
        {
            var trace = current; trace?.Record(phase, "Entered", null);
            try
            {
                T value = action(); trace?.Record(phase, "Returned", null); return value;
            }
            catch (Exception error) { trace?.Record(phase, "Faulted", error); throw; }
        }
        internal static void Run(Phase phase, Action action)
        { Read(phase, () => { action(); return true; }); }

        private void Record(Phase phase, string boundary, Exception error)
        {
            if (writer == null || disposed || events >= MaximumEvents || !Enum.IsDefined(typeof(Phase), phase)) return;
            events++;
            try
            {
                writer(new JavaScriptSerializer().Serialize(new
                {
                    Correlation = correlation, Sequence = events, Phase = phase.ToString(), Boundary = boundary,
                    ElapsedTicks = clock.ElapsedTicks, ClockFrequency = Stopwatch.Frequency,
                    ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString(),
                    ContextType = SynchronizationContext.Current?.GetType().FullName,
                    HResult = error == null ? (int?)null : error.HResult, ObservedUtc = DateTime.UtcNow.ToString("o")
                }));
            }
            catch { /* Deliberately passive, including serialization and file failures. */ }
        }
        public void Dispose()
        {
            if (disposed) return;
            Record(Phase.Scope, "Returned", null); disposed = true; current = previous;
        }
    }
}
