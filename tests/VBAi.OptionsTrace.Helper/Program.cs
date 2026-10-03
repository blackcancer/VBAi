using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.OptionsTrace
{
    // A disposable CLR/SOS preflight. Never creates an Office application or opens native Options.
    internal static class Program
    {
        private static int Main()
        {
            using (var own = Process.GetCurrentProcess())
                Console.WriteLine("READY " + own.Id + " " + own.StartTime.ToUniversalTime().ToString("o"));
            var capture = typeof(VbeDebugWindows).GetMethod("CaptureOptionsTabs", BindingFlags.Static | BindingFlags.NonPublic);
            var expected = capture.Invoke(null, new object[] { new SyntheticOptions(), new IntPtr(1) });
            Console.WriteLine("EXPECTED " + Convert.ToBase64String(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(expected))));
            if (Console.ReadLine() != "OBSERVE") return 2;
            try { RejectStaleRevision(); return 3; }
            catch (InvalidOperationException error)
            {
                if (error.Message != "VBE options changed since inspection; read them again.") return 4;
                Console.WriteLine("GUARD_REJECTED");
            }
            return Console.ReadLine() == "QUIT" ? 0 : 5;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void RejectStaleRevision()
        {
            VbeDebugWindows.SetVbeOption(new Request {
                Pane = "Editor Format", Property = "Font", Value = "Synthetic alternate",
                ExpectedOptionsVersion = new string('0', 64)
            }, new SyntheticOptions());
        }

        private sealed class SyntheticOptions : VbeDebugWindows.IWritableOptionsProbe, VbeDebugWindows.IFormatCategoriesOptionsProbe
        {
            public IntPtr Dialog() => new IntPtr(1);
            public IList<string> Tabs(IntPtr dialog) => new[] { "Editor Format" };
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int tab) => new[] {
                new VbeDebugWindows.OptionsControl { Name = "Font", Type = "ControlType.ComboBox",
                    Value = "Synthetic original", Visible = true, Enabled = true,
                    Choices = Enumerable.Range(0, 700).Select(index => "Synthetic font " + index + " (Occidental)").ToArray(),
                    NativeChoices = Enumerable.Range(0, 700).Select(index => new VbeDebugWindows.OptionsNativeChoice {
                        Index = index, Label = "Synthetic font " + index + " (Occidental)", SelectionValue = "Synthetic font " + index + " (Occidental)" }).ToArray() },
                new VbeDebugWindows.OptionsControl { Name = "Synthetic checkbox", Type = "ControlType.RadioButton",
                    Value = true, Visible = true, Enabled = true }
            };
            public IList<VbeDebugWindows.OptionsFormatCategory> FormatCategories(IntPtr dialog, int tab) => new[] {
                new VbeDebugWindows.OptionsFormatCategory { Category = "Texte normal", Palettes = new[] {
                    new VbeDebugWindows.OptionsControl { Name = "Premier plan :", Type = "ControlType.ComboBox",
                        Value = " Automatique", Visible = true, Enabled = false, SelectedIndex = 0,
                        Choices = new[] { " Automatique", "NativeIndex:1" }, NativeChoices = new[] {
                            new VbeDebugWindows.OptionsNativeChoice { Index = 0, Label = " Automatique", SelectionValue = " Automatique" },
                            new VbeDebugWindows.OptionsNativeChoice { Index = 1, Label = "", SelectionValue = "NativeIndex:1" } } }
                } }
            };
            public void SelectFormatCategory(IntPtr dialog, int tab, string category)
            { throw new InvalidOperationException("A stale guard must refuse before category selection."); }
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Pause(int milliseconds) { }
            public void Close(IntPtr dialog) { }
            public void Write(IntPtr dialog, int tab, string name, string type, object value)
            { throw new InvalidOperationException("The stale guard must refuse before any write."); }
            public void Accept(IntPtr dialog)
            { throw new InvalidOperationException("The stale guard must refuse before Accept."); }
        }
    }
}
