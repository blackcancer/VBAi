namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeSessionContractTests
    {
        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }

        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
        }

        public sealed class FakeModule
        {
            private readonly List<string> lines;
            public FakeModule() : this(0)
            {
            }

            public FakeModule(int count)
            {
                lines = Enumerable.Repeat(string.Empty, count).ToList();
                Lines = new FakeLines(this);
            }

            public FakeModule(string code)
            {
                lines = code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                Lines = new FakeLines(this);
            }

            public int CountOfLines
            {
                get
                {
                    return lines.Count;
                }

                set
                {
                    lines.Clear();
                    lines.AddRange(Enumerable.Repeat(string.Empty, value));
                }
            }

            public string Code
            {
                get
                {
                    return string.Join("\r\n", lines);
                }
            }

            public FakeLines Lines { get; }
            public bool CorruptNonAsciiOnInsert { get; set; }

            public void DeleteLines(int start, int count)
            {
                lines.RemoveRange(start - 1, count);
            }

            public void InsertLines(int start, string text)
            {
                if (CorruptNonAsciiOnInsert)
                    text = text.Replace('é', '?');
                lines.InsertRange(start - 1, text.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }

            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module)
                {
                    this.module = module;
                }

                public string this[int start, int count]
                {
                    get
                    {
                        return string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
                    }
                }
            }
        }
    }
}
