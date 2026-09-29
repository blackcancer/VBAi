namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsOptionsProbeTests
    {
        private static OptionsFake DebugFake()
        {
            var fake = new OptionsFake();
            fake.Choices.Add(new VbeDebugWindows.OptionsChoice { Name = "Break on All Errors", Selected = true });
            fake.Choices.Add(new VbeDebugWindows.OptionsChoice { Name = "Break in Class Module" });
            fake.Choices.Add(new VbeDebugWindows.OptionsChoice { Name = "Break on Unhandled Errors" });
            return fake;
        }

        private sealed class OptionsFake : VbeDebugWindows.IOptionsProbe
        {
            public readonly List<string> Names = new List<string>
            {
                "Editor"
            };
            public readonly List<VbeDebugWindows.OptionsControl> Items = new List<VbeDebugWindows.OptionsControl>();
            public readonly List<VbeDebugWindows.OptionsChoice> Choices = new List<VbeDebugWindows.OptionsChoice>();
            public bool Open = true, CloseAfterRead;
            public int Closes, Pauses, ClosePolls;
            public IntPtr Dialog()
            {
                if (Closes > 0)
                    ClosePolls++;
                return Open && !(Closes > 0 && CloseAfterRead) ? new IntPtr(2) : IntPtr.Zero;
            }

            public IList<string> Tabs(IntPtr dialog)
            {
                return Names;
            }

            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int tabIndex)
            {
                return Items;
            }

            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog)
            {
                return Choices;
            }

            public void Close(IntPtr dialog)
            {
                Closes++;
            }

            public void Pause(int milliseconds)
            {
                Pauses++;
            }
        }
    }
}
