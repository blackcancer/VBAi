using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class ModernEditorThemeTests
    {
        [STATestMethod]
        public void WindowsWorkerThemeNotificationUpdatesMonacoAndTabsOnTheUiThread()
        {
            using (var theme = new ThemeScope())
            using (var f = new ModernEditorDebugFixture())
            {
                int owner = Thread.CurrentThread.ManagedThreadId;
                foreach (var choice in new[] { ThemeChoice.Light, ThemeChoice.Dark, ThemeChoice.Light })
                {
                    ThemeScope.SetChoice(choice);
                    var observed = new TaskCompletionSource<bool>();
                    f.Window.ScriptExecution = (method, values) =>
                    {
                        if (method == "theme")
                        {
                            try
                            {
                                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                                Assert.AreEqual(choice == ThemeChoice.Dark, values[0]);
                                Assert.AreEqual(UiTheme.Surface.ToArgb(), ColorTranslator.FromHtml((string)values[2]).ToArgb());
                                StringAssert.StartsWith((string)values[2], "#");
                                using (var bitmap = new Bitmap(500, 60))
                                using (var graphics = Graphics.FromImage(bitmap))
                                {
                                    UiInvoke.Call(typeof(ThemedTabControl), "OnPaint", f.Get<ThemedTabControl>("tabs"),
                                        new System.Windows.Forms.PaintEventArgs(graphics, new Rectangle(0, 0, 500, 60)));
                                    Assert.AreEqual(UiTheme.Background.ToArgb(), bitmap.GetPixel(499, 59).ToArgb());
                                }
                                observed.SetResult(true);
                            }
                            catch (Exception error) { observed.TrySetException(error); }
                        }
                        return Task.FromResult("null");
                    };
                    ModernEditorDebugFixture.Wait(Task.Run(() => UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", f.Window)));
                    ModernEditorDebugFixture.Wait(observed.Task);
                }
            }
        }
    }
}
