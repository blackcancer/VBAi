namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeMenuLifecycleTests
    {
        private static FakeHost Host()
        {
            var view = new FakeButton
            {
                Caption = "&View"
            };
            var tools = new FakeButton
            {
                Caption = "&Tools"
            };
            var main = new FakeBar
            {
                Type = 1,
                Controls = new FakeControls()
            };
            main.Controls.Items.Add(view);
            main.Controls.Items.Add(tools);
            return new FakeHost
            {
                CommandBars = new[]
                {
                    main,
                    new FakeBar
                    {
                        Type = 0,
                        Name = "Code Window",
                        Controls = new FakeControls()
                    }
                }
            };
        }

        public sealed class FakeHost
        {
            public FakeBar[] CommandBars { get; set; }
        }

        public sealed class FakeBar
        {
            public int Type { get; set; }
            public string Name { get; set; }
            public FakeControls Controls { get; set; }
        }

        public sealed class FakeButton
        {
            public string Caption { get; set; }
            public string Tag { get; set; }
            public string TooltipText { get; set; }
            public object Picture { get; set; }
            public object Mask { get; set; }
            public int Style { get; set; }
            public bool RejectDelete { get; set; }
            public FakeControls Controls { get; } = new FakeControls();
            public int DeleteCount { get; private set; }

            public void Delete()
            {
                if (RejectDelete) throw new InvalidOperationException("Delete rejected");
                DeleteCount++;
            }
        }

        public sealed class FakeControls : IEnumerable<FakeButton>
        {
            public List<FakeButton> Items { get; } = new List<FakeButton>();
            public int FailAtCount { get; set; } = int.MaxValue;

            public FakeButton Add(int type, object id, object parameter, object before, bool temporary)
            {
                if (Items.Count >= FailAtCount)
                    throw new InvalidOperationException("Add failed");
                var button = new FakeButton();
                Items.Add(button);
                return button;
            }

            public IEnumerator<FakeButton> GetEnumerator()
            {
                return Items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
