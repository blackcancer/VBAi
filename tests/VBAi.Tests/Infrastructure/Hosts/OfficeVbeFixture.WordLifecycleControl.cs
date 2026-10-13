using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Text;
using System.Web.Script.Serialization;
using System.IO.Packaging;
using System.Security.Cryptography;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private string wordLifecycleTemplate, wordLifecycleTemplateHash;

        private string PrepareWordLifecycleTemplate()
        {
            if (wordLifecycleTemplate != null) return wordLifecycleTemplate;
            RequireCoverageConsentOwner();
            Assert.AreEqual("Word", Kind);
            string path = Path.Combine(Root, "WordLifecycleTemplate.dotx");
            WriteWordLifecycleTemplateSeed(path);
            wordLifecycleTemplateHash = HashOwnedFile(path);
            wordLifecycleTemplate = path;
            return path;
        }

        /// <summary>Creates one fresh minimal DOTX without VBA, external relationships or personal-template dependencies.</summary>
        internal static void WriteWordLifecycleTemplateSeed(string path)
        {
            WriteMacroFreeSeed(path);
            using (var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite))
            {
                var uri = new Uri("/word/document.xml", UriKind.Relative);
                byte[] body;
                using (var stream = package.GetPart(uri).GetStream())
                using (var bytes = new MemoryStream()) { stream.CopyTo(bytes); body = bytes.ToArray(); }
                package.DeletePart(uri);
                var part = package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml");
                using (var stream = part.GetStream()) stream.Write(body, 0, body.Length);
            }
        }

        private void WriteWordLifecycleControlSeed(string path)
        {
            WriteMacroFreeSeed(path);
            using (var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite))
            {
                var settings = package.CreatePart(new Uri("/word/settings.xml", UriKind.Relative),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
                const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<w:attachedTemplate r:id=\"rTemplate\"/></w:settings>";
                using (var writer = new StreamWriter(settings.GetStream(), new UTF8Encoding(false))) writer.Write(xml);
                settings.CreateRelationship(new Uri(Path.Combine(Root, "WordLifecycleTemplate.dotx")), TargetMode.External,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", "rTemplate");
                package.GetPart(new Uri("/word/document.xml", UriKind.Relative)).CreateRelationship(settings.Uri, TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
            }
        }
        /// <summary>Calibrates visible/hidden document events and auto-macro suppression on fresh macro-free controls only.</summary>
        internal object[] CalibrateWordLifecycle()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_LIFECYCLE_CALIBRATION") != "1") return null;
            RequireCoverageConsentOwner();
            Assert.AreEqual("Word", Kind);
            Assert.IsFalse(NativeExecutionUnsettled);
            RequireMainWordOriginal(); // Requires the original /a macro-free launch, not a generic existing Word instance.
            string sourceHash = ReadWordLifecycleSourceHash(DocumentPath);
            var results = new System.Collections.Generic.List<object>();
            foreach (bool suppress in new[] { true, false })
            foreach (bool visible in new[] { false, true })
            {
                string path = Path.Combine(Root, "lifecycle-control-" + (suppress ? "suppressed" : "unsuppressed") + (visible ? "-visible.docx" : "-hidden.docx"));
                WriteWordLifecycleControlSeed(path);
                string hash = HashOwnedFile(path);
                var control = new OfficeWordLifecycleControl();
                var templateReviews = new System.Collections.Generic.List<object>();
                Exception failure = null;
                var done = new ManualResetEventSlim();
                Thread worker = null;
                SaveWordLifecycleReceipt(path, "intent", new { ProcessId, DocumentPath, Path = path, Sha256 = hash,
                    OriginalSha256 = sourceHash, Visible = visible, SuppressedAutoMacros = suppress, MacroSecurityRequired = 3, DiagnosticOnly = true });
                try
                {
                    using (var observer = new OfficeCoverageEventObserver(this))
                    {
                        NativeExecutionUnsettled = true;
                        // The worker acquires its own application lease on its own STA. It never
                        // receives the fixture's COM objects. The original owner pumps its event reads.
                        worker = new Thread(() =>
                        {
                            try
                            {
                                var boundary = new OfficeWordLifecycleControl.Boundary
                                {
                                    Acquire = () => Marshal.GetActiveObject("Word.Application"),
                                    RequireOwner = value =>
                                    {
                                        RequireMainWordOriginal();
                                        IntPtr window = VbaTestWordValuesHost.ReadApplicationWindow(value);
                                        uint pid; Assert.AreNotEqual(0u, GetWindowThreadProcessId(window, out pid));
                                        Assert.AreEqual((uint)ProcessId, pid);
                                    },
                                    RequireCopy = () => RequireUnchangedOpenWordSeed(path, hash),
                                    RequireAutoMacroScope = value => RequireReviewedWordControlTemplates(value, path, hash, templateReviews.Add),
                                    Paths = ReadWordLifecyclePaths,
                                    ReadSecurity = value => Convert.ToInt32(((dynamic)value).AutomationSecurity),
                                    WriteSecurity = (value, security) => ((dynamic)value).AutomationSecurity = security,
                                    WordBasic = value => ((dynamic)value).WordBasic,
                                    SuppressAutoMacros = (value, state) => ((dynamic)value).DisableAutoMacros(state),
                                    Open = (value, copy, show) =>
                                    {
                                        var lease = new VbaTestComReferences();
                                        var returned = new VbaTestComReferences();
                                        try
                                        {
                                            object documents = lease.Add((object)((dynamic)value).Documents);
                                            object opened = returned.Add((object)((dynamic)documents).Open(FileName: copy, ConfirmConversions: false,
                                                ReadOnly: true, AddToRecentFiles: false, Revert: false, Visible: show, OpenAndRepair: false));
                                            lease.Dispose();
                                            returned.Transfer(opened);
                                            returned.Dispose();
                                            return opened;
                                        }
                                        catch { lease.Retain(); returned.Retain(); throw; }
                                    },
                                    DocumentPath = value => Convert.ToString(((dynamic)value).FullName),
                                    HasVba = value => Convert.ToBoolean(((dynamic)value).HasVBProject),
                                    ReadOnly = value => Convert.ToBoolean(((dynamic)value).ReadOnly),
                                    Close = value => ((dynamic)value).Close(0)
                                };
                                control.Run(boundary, DocumentPath, path, visible, suppress);
                            }
                            catch (Exception error) { failure = error; }
                            finally { done.Set(); }
                        }) { IsBackground = true };
                        worker.SetApartmentState(ApartmentState.STA);
                        worker.Start();
                        try
                        {
                            OfficeCoverageEventObserver.ReadPumping(() =>
                            {
                                if (!done.Wait(60000)) throw new TimeoutException("Original Word control is still live; do not repeat or clean up uncertain native work.");
                                return true;
                            }, Application.DoEvents, Thread.Sleep);
                            RequireUnchangedOpenWordSeed(DocumentPath, sourceHash);
                        }
                        catch (Exception observation)
                        {
                            SaveWordLifecycleReceipt(path, "observation-failure", new { Error = observation.ToString(),
                                NativeExecutionUnsettled, WorkerStillRunning = worker.IsAlive,
                                Trace = SnapshotWordLifecycleTrace(observer.Trace), NoReplay = true, DiagnosticOnly = true });
                            throw;
                        }
                        // Only a known terminal refusal or fully released success permits fixture cleanup.
                        NativeExecutionUnsettled = control.Uncertain;
                        var trace = observer.Trace;
                        var result = new { Path = path, Visible = visible, SuppressedAutoMacros = suppress,
                            control.OpenEnteredWithForceDisable, control.OriginalSuppressionRestoreReturned, control.OriginalSecurityRestored,
                            Trace = SnapshotWordLifecycleTrace(trace), Error = failure?.ToString(),
                            control.Uncertain, TemplateReviews = templateReviews.ToArray(), DiagnosticOnly = true };
                        results.Add(result);
                        SaveWordLifecycleReceipt(path, "terminal", result);
                        if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
                    }
                }
                finally
                {
                    // A timed-out native worker may still signal completion. Never dispose its
                    // signal underneath it or let that secondary exception kill the test process.
                    if (worker == null || !worker.IsAlive) done.Dispose();
                }
            }
            return results.ToArray();
        }

        /// <summary>Reads the owned open source without denying Word's existing write handle or changing its bytes.</summary>
        internal static string ReadWordLifecycleSourceHash(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
                throw new ArgumentException("An absolute owned Word source path is required.");
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
        }

        private static object SnapshotWordLifecycleTrace(OfficeCoverageEventTrace trace)
        {
            lock (trace.Events)
            lock (trace.Failures)
                return new { trace.OwnerThread, Events = trace.Events.ToArray(), Failures = trace.Failures.ToArray() };
        }

        /// <summary>Requires generated templates or the independently reviewed empty Normal object in the exact owned /a process.</summary>
        private void RequireReviewedWordControlTemplates(object app, string copy, string copyHash, Action<object> recordMetadata)
        {
            RequireReviewedWordEventTemplates(app, recordMetadata);
            string expected = Path.GetFullPath(Path.Combine(Root, "WordLifecycleTemplate.dotx"));
            InspectWordControlPackage(copy, copyHash, package =>
            {
                var settings = package.GetPart(new Uri("/word/settings.xml", UriKind.Relative));
                var attached = settings.GetRelationship("rTemplate");
                Assert.AreEqual(TargetMode.External, attached.TargetMode);
                Assert.AreEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", attached.RelationshipType);
                Assert.AreEqual(new Uri(expected), attached.TargetUri, "The control must not fall back to a personal Normal template.");
            });
        }

        private void RequireReviewedWordEventTemplates(object app, Action<object> recordMetadata)
        {
            RequireMainWordOriginal();
            string expected = Path.GetFullPath(Path.Combine(Root, "WordLifecycleTemplate.dotx"));
            Assert.AreEqual(expected, wordLifecycleTemplate, true);
            RequireUnchangedOpenWordSeed(expected, wordLifecycleTemplateHash);
            using (var references = new VbaTestComReferences())
            {
                var documents = references.Items((object)((dynamic)app).Documents).Take(4).ToArray();
                Assert.AreEqual(1, documents.Length);
                Assert.AreEqual(DocumentPath, (string)((dynamic)documents[0]).FullName, true);
                object attached = references.Add((object)((dynamic)documents[0]).AttachedTemplate);
                Assert.AreEqual(expected, (string)((dynamic)attached).FullName, true,
                    "The source must remain attached to its generated macro-free template.");
                var templates = references.Items((object)((dynamic)app).Templates).Take(33).ToArray();
                Assert.IsTrue(templates.Length < 33, "The loaded template inventory is incomplete.");
                object normal = null;
                foreach (object template in templates)
                {
                    if (string.Equals(expected, Convert.ToString(((dynamic)template).FullName), StringComparison.OrdinalIgnoreCase)) continue;
                    if (normal == null) normal = references.Add((object)((dynamic)app).NormalTemplate);
                    recordMetadata(ReviewEmptyStartupNormalTemplate(template, normal));
                }
            }
            Assert.IsTrue(File.Exists(expected));
            Assert.AreEqual(0, (int)(File.GetAttributes(expected) & FileAttributes.ReparsePoint));
            InspectWordControlPackage(expected, wordLifecycleTemplateHash, package =>
            {
                var parts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                Assert.AreEqual(1, parts.Length);
                Assert.AreEqual("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", parts[0].ContentType);
                Assert.AreEqual(0, parts[0].GetRelationships().Count());
                var links = package.GetRelationships().ToArray();
                Assert.AreEqual(1, links.Length);
                Assert.AreEqual(TargetMode.Internal, links[0].TargetMode);
                Assert.AreEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", links[0].RelationshipType);
                Assert.AreEqual(parts[0].Uri, links[0].TargetUri);
            });
        }

        private void SaveWordLifecycleReceipt(string copy, string phase, object value)
        {
            string receipt = Path.Combine(Root, Path.GetFileNameWithoutExtension(copy) + "-" + phase + ".json");
            using (var stream = new FileStream(receipt, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                writer.Write(new JavaScriptSerializer().Serialize(value));
        }

        private static string[] ReadWordLifecyclePaths(object application)
        {
            using (var references = new VbaTestComReferences())
            {
                object documents = (object)((dynamic)application).Documents;
                return references.Items(documents).Select(value => (string)Convert.ToString(((dynamic)value).FullName)).Take(4).ToArray();
            }
        }
    }
}
