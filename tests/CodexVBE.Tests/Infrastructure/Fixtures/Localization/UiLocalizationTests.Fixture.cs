namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class UiLocalizationTests
    {
        private static FakeVbe Host(string caption)
        {
            return new FakeVbe
            {
                CommandBars = new[]
                {
                    new FakeBar
                    {
                        Type = 1,
                        Controls = new[]
                        {
                            new FakeControl
                            {
                                Caption = caption
                            }
                        }
                    }
                }
            };
        }

        public sealed class FakeVbe
        {
            public FakeBar[] CommandBars { get; set; }
        }

        public sealed class FakeBar
        {
            public int Type { get; set; }
            public FakeControl[] Controls { get; set; }
        }

        public sealed class FakeControl
        {
            public string Caption { get; set; }
        }
    }
}
