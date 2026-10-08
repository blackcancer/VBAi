using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises complete transcript workloads without showing windows or using a live provider.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Records native append and full-replacement costs for identical streamed snapshots.</summary>
        [STATestMethod, TestCategory("Performance")]
        [DataRow(2048, 340)]
        [DataRow(32768, 340)]
        [DataRow(131072, 340)]
        [DataRow(2048, 780)]
        [DataRow(32768, 780)]
        [DataRow(131072, 780)]
        public void StreamedNativeTextWorkloadRetainsSelectionAndReportsFrameCosts(int characters, int width)
        {
            using (var incremental = new ChatTextContentView())
            using (var replacement = new ChatTextContentView())
            {
                foreach (var view in new[] { incremental, replacement })
                {
                    view.Width = width;
                    var handle = view.content.Handle;
                    view.ShowPlain("Start");
                    view.content.Select(1, 3);
                }
                int changes = 0;
                incremental.content.TextChanged += (s, e) => changes++;
                const int frames = 64;
                var snapshots = Enumerable.Range(1, frames).Select(i => "Start" + new string('x', characters * i / frames)).ToArray();
                var baseline = MeasureTranscriptFrames(snapshots, text =>
                {
                    int start = replacement.content.SelectionStart, length = replacement.content.SelectionLength;
                    replacement.content.Text = text;
                    replacement.content.Select(start, length);
                });
                var optimized = MeasureTranscriptFrames(snapshots, text => incremental.UpdatePlainText(text));
                Assert.AreEqual(snapshots.Last(), incremental.content.Text);
                Assert.AreEqual(replacement.content.Text, incremental.content.Text);
                Assert.AreEqual(frames, changes, "Each frame must publish exactly once.");
                Assert.AreEqual(1, incremental.content.SelectionStart);
                Assert.AreEqual(3, incremental.content.SelectionLength);
                Assert.IsTrue(incremental.content.Height >= 24 && incremental.content.Height <= 1200);
                ReportTranscriptTimings("native-append", characters, width, optimized);
                ReportTranscriptTimings("native-full-replacement", characters, width, baseline);
            }
        }

        /// <summary>Exercises capped multiline height through repeated streaming and narrow/wide layout transitions.</summary>
        [STATestMethod, TestCategory("Performance")]
        public void MultilineStreamingAndResizeRetainTheCappedViewport()
        {
            using (var view = new ChatTextContentView())
            {
                view.Width = 400;
                var handle = view.content.Handle;
                string text = string.Concat(Enumerable.Repeat("Line of streamed transcript text.\n", 200));
                view.ShowPlain(text);
                Assert.AreEqual(1200, view.content.Height);
                for (int i = 0; i < 20; i++)
                {
                    text += "Another streamed line.\n";
                    view.UpdatePlainText(text);
                    Assert.AreEqual(1200, view.content.Height, "Streaming must not collapse an internally scrolled transcript.");
                }
                foreach (int width in new[] { 220, 780, 340, 600 })
                {
                    view.Width = width;
                    view.PerformLayout();
                    Assert.AreEqual(text, view.content.Text);
                    Assert.AreEqual(1200, view.content.Height);
                }
                view.UpdatePlainText("Short result");
                Assert.IsTrue(view.content.Height < 80);
                Assert.AreEqual(System.Windows.Forms.RichTextBoxScrollBars.None, view.content.ScrollBars);
            }
        }

        /// <summary>Measures paged histories followed by a streamed reply without recycling earlier realized entries.</summary>
        [STATestMethod, TestCategory("Performance")]
        [DataRow(80)]
        [DataRow(500)]
        [DataRow(2000)]
        public void LargeHistoryStreamingKeepsEarlierViewsAndCoalescesFragments(int historyCount)
        {
            using (var window = Surfaces())
            {
                Set(window, "loadingSession", true);
                for (int i = 0; i < historyCount; i++)
                    Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "History entry " + i });
                Set(window, "loadingSession", false);
                var timer = Stopwatch.StartNew();
                Call(window, "RefreshTranscriptWindow", Math.Max(0, historyCount - 80));
                double projectionMs = timer.Elapsed.TotalMilliseconds;
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                var earlier = entries[historyCount - 1];
                var previousItem = new TranscriptItem { DataContext = earlier };
                Call(window, "RealizeEntry", previousItem);
                using (var previous = (ChatDesignerHost)previousItem.Content)
                {
                    Call(window, "ReceiveChatUpdate", "final", "performance-stream", "", false);
                    var stream = entries.Last();
                    using (var live = (ChatDesignerHost)Call(window, "RenderEntry", stream))
                    {
                        int collections = 0, textChanges = 0;
                        var visible = Get<ObservableCollection<object>>(window, "visibleEntries");
                        visible.CollectionChanged += (s, e) => collections++;
                        var text = Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts")["performance-stream"];
                        var handle = text.Handle;
                        text.TextChanged += (s, e) => textChanges++;
                        timer.Restart();
                        for (int i = 0; i < 2000; i++) Call(window, "ReceiveChatUpdate", "final", "performance-stream", "x", false);
                        Assert.AreEqual(0, textChanges);
                        Call(window, "FlushStreamText");
                        double streamMs = timer.Elapsed.TotalMilliseconds;
                        Assert.AreEqual(new string('x', 2000), text.Text);
                        Assert.AreEqual(1, textChanges);
                        Assert.AreEqual(0, collections, "Streaming must not rebuild the visible collection.");
                        Assert.AreSame(previous, Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews")[earlier]);
                        Assert.AreEqual("History entry " + (historyCount - 1), earlier.Text);
                        Console.WriteLine("TRANSCRIPT history={0} projection_ms={1:F3} fragments_and_flush_ms={2:F3}", historyCount, projectionMs, streamMs);
                    }
                }
            }
        }

        /// <summary>Measures a live activity among retained completed siblings without rebuilding their controls.</summary>
        [STATestMethod, TestCategory("Performance")]
        [DataRow(20, false)]
        [DataRow(20, true)]
        [DataRow(100, false)]
        [DataRow(100, true)]
        public void ActivityWorkloadKeepsSiblingControlsAndOnePublicationPerFrame(int siblings, bool expanded)
        {
            using (var window = Surfaces())
            {
                for (int i = 0; i < siblings; i++)
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "saved-" + i, Kind = "dynamicToolCall", Title = "Read module", Detail = "Saved result " + i, Status = "completed" });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "live", Kind = "reasoning", Title = "Reasoning", Detail = "Start", Status = "inProgress" });
                var owner = Get<List<ChatEntry>>(window, "transcriptEntries")[0];
                var item = new TranscriptItem { DataContext = owner };
                Call(window, "RealizeEntry", item);
                using (var host = (ChatDesignerHost)item.Content)
                {
                    var group = (ChatActivityGroupView)host.View;
                    group.section.Expanded = expanded;
                    var controls = group.section.body.Controls.OfType<ChatActivityStepView>().ToArray();
                    Assert.AreEqual(siblings + 1, controls.Length);
                    var text = Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts")["live"];
                    var handle = text.Handle;
                    text.Select(1, 3);
                    int changes = 0;
                    text.TextChanged += (s, e) => changes++;
                    var times = new List<double>();
                    for (int frame = 0; frame < 32; frame++)
                    {
                        var timer = Stopwatch.StartNew();
                        for (int fragment = 0; fragment < 20; fragment++)
                            Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "live", Kind = "reasoning", Detail = ".", Status = "inProgress", Append = true });
                        Call(window, "FlushStreamText");
                        times.Add(timer.Elapsed.TotalMilliseconds);
                    }
                    Assert.AreEqual(32, changes);
                    Assert.AreEqual("Start" + new string('.', 640), text.Text);
                    Assert.AreEqual("tar", text.SelectedText);
                    for (int i = 0; i < controls.Length; i++)
                    {
                        Assert.AreSame(controls[i], group.section.body.Controls[i]);
                        Assert.IsFalse(controls[i].IsDisposed);
                    }
                    ReportTranscriptTimings("activity-" + (expanded ? "expanded" : "collapsed"), siblings, 0, times.ToArray());
                }
            }
        }

        /// <summary>Measures each UI frame without imposing a machine-dependent speed threshold.</summary>
        private static double[] MeasureTranscriptFrames(string[] snapshots, Action<string> publish)
        {
            return snapshots.Select(text => { var timer = Stopwatch.StartNew(); publish(text); return timer.Elapsed.TotalMilliseconds; }).ToArray();
        }

        /// <summary>Writes reproducible workload dimensions and distribution statistics to the retained TRX output.</summary>
        private static void ReportTranscriptTimings(string workload, int size, int width, double[] times)
        {
            var sorted = times.OrderBy(value => value).ToArray();
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "TRANSCRIPT workload={0} size={1} width={2} frames={3} total_ms={4:F3} median_ms={5:F3} p95_ms={6:F3} max_ms={7:F3}",
                workload, size, width, times.Length, times.Sum(), sorted[sorted.Length / 2], sorted[(int)Math.Ceiling(sorted.Length * 0.95) - 1], sorted.Last()));
        }
    }
}
