using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeDebugWindowsObserverTextTests
    {
        [TestMethod]
        public void HungObserverTitleHasOneBoundedReadNoActionAndRestoresLegacyReadScope()
        {
            using (var scene = new NativeDebugScene())
            {
                scene.Add("#32770", "Quick Watch");
                int boundedReads = 0, legacyReads = 0;
                VbeDebugWindows.ReadObserverTextMessage = (IntPtr handle, uint message, IntPtr capacity, StringBuilder text,
                    uint flags, uint milliseconds, out UIntPtr result) =>
                {
                    boundedReads++;
                    Assert.AreEqual(13u, message);
                    Assert.AreEqual(3u, flags);
                    Assert.AreEqual(500u, milliseconds);
                    Assert.AreEqual(512, capacity.ToInt32());
                    result = UIntPtr.Zero;
                    return IntPtr.Zero;
                };
                VbeDebugWindows.GetWindowText = (handle, text, capacity) => { legacyReads++; text.Append(scene.Find(handle).Caption); return text.Length; };
                var rows = new List<string>();
                var trace = new VbeInspectionTrace(rows.Add);
                using (trace.Enter()) Assert.ThrowsException<TimeoutException>(() => VbeDebugWindows.ReadScalarQuickWatch(new Request { Expression = "secret expression" }));
                Assert.AreEqual(1, boundedReads);
                Assert.AreEqual(0, legacyReads);
                Assert.AreEqual(0, scene.Messages.Count);
                StringAssert.Contains(rows[0], "ObserverEntered");
                StringAssert.Contains(rows[1], "ObserverTerminal");
                StringAssert.Contains(rows[1], "TimeoutException");
                Assert.IsFalse(string.Join("", rows).Contains("secret expression"));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoQuickWatchDialog());
                Assert.AreEqual(1, boundedReads);
                Assert.IsTrue(legacyReads > 0);
            }
        }

        [TestMethod]
        public void SuccessfulEmptyTextIsNotTimeoutAndTruncatedTextIsNotAccepted()
        {
            foreach (bool truncated in new[] { false, true })
                using (var scene = new NativeDebugScene())
                {
                    var dialog = scene.Add("#32770", "Quick Watch");
                    var expression = scene.Add("Edit", "x", dialog);
                    var value = scene.Add("Static", "", dialog);
                    var context = scene.Add("Edit", "P.M.Run", dialog);
                    var cancel = scene.Add("Button", "Cancel", dialog);
                    var oldItem = VbeDebugWindows.GetDlgItem;
                    var oldPost = VbeDebugWindows.PostMessage;
                    var oldPause = VbeDebugWindows.PauseNative;
                    try
                    {
                        int closes = 0;
                        VbeDebugWindows.GetDlgItem = (handle, id) => handle != dialog.Handle ? IntPtr.Zero : id == 4751 ? expression.Handle : id == 4752 ? value.Handle : id == 4753 ? context.Handle : id == 2 ? cancel.Handle : IntPtr.Zero;
                        VbeDebugWindows.PauseNative = _ => { };
                        VbeDebugWindows.PostMessage = (handle, message, w, l) => { Assert.AreEqual(cancel.Handle, handle); closes++; dialog.Visible = false; return true; };
                        VbeDebugWindows.ReadObserverTextMessage = (IntPtr handle, uint message, IntPtr capacity, StringBuilder text,
                            uint flags, uint milliseconds, out UIntPtr result) =>
                        {
                            text.Append(scene.Find(handle).Caption);
                            result = new UIntPtr(handle == value.Handle && truncated ? 511u : (uint)text.Length);
                            return new IntPtr(1);
                        };
                        var request = new Request { Project = "P", Module = "M", Procedure = "Run", Expression = "x" };
                        if (truncated)
                            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadScalarQuickWatch(request)).Message, "truncated");
                        else
                        {
                            dynamic observed = VbeDebugWindows.ReadScalarQuickWatch(request);
                            Assert.AreEqual("", (string)observed.Value);
                        }
                        Assert.AreEqual(1, closes);
                        Assert.IsFalse(dialog.Visible);
                    }
                    finally { VbeDebugWindows.GetDlgItem = oldItem; VbeDebugWindows.PostMessage = oldPost; VbeDebugWindows.PauseNative = oldPause; }
                }
        }
    }
}
