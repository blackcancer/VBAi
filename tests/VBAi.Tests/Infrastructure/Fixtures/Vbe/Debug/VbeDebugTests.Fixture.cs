namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugTests
    {
        private const string Code = "Debug.Print 1\r\nDebug.Print 2";
        private static Fixture Create(int mode = 2)
        {
            var project = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Debug.xlsm",
                Mode = mode
            };
            var component = new FakeComponent
            {
                Name = "Module1"
            };
            var module = new FakeModule(component, Code);
            component.CodeModule = module;
            project.VBComponents.Add(component);
            var vbe = new FakeVbe
            {
                ActiveVBProject = project,
                ActiveCodePane = module.CodePane
            };
            vbe.VBProjects.Add(project);
            var bar = new FakeBar
            {
                Name = "Debug"
            };
            vbe.CommandBars.Add(bar);
            return new Fixture
            {
                Vbe = vbe,
                Project = project,
                Module = module,
                Bar = bar,
                Service = new VbeDebug(vbe)
            };
        }

        private static Request Location(Fixture f, int line = 1)
        {
            return new Request
            {
                Project = f.Project.Name,
                Module = "Module1",
                StartLine = line,
                ExpectedSha256 = Sha(Code),
                ExpectedMode = f.Project.Mode
            };
        }

        private static string Sha(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant();
        }

        private sealed class RecordingContext : SynchronizationContext
        {
            private readonly List<Tuple<SendOrPostCallback, object>> callbacks = new List<Tuple<SendOrPostCallback, object>>();
            public int Count => callbacks.Count;

            public override void Post(SendOrPostCallback callback, object state)
            {
                callbacks.Add(Tuple.Create(callback, state));
            }

            public void RunAll()
            {
                foreach (var callback in callbacks)
                    callback.Item1(callback.Item2);
                callbacks.Clear();
            }
        }

        private sealed class Fixture
        {
            public FakeVbe Vbe;
            public FakeProject Project;
            public FakeModule Module;
            public FakeBar Bar;
            public VbeDebug Service;
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
            public List<FakeBar> CommandBars { get; } = new List<FakeBar>();
            public FakeProject ActiveVBProject { get; set; }
            private FakePane activeCodePane;
            public bool IgnoreActiveCodePaneAssignment { get; set; }
            public int ActiveCodePaneSetCount { get; private set; }
            public FakePane ActiveCodePane
            {
                get => activeCodePane;
                set
                {
                    ActiveCodePaneSetCount++;
                    if (!IgnoreActiveCodePaneAssignment) activeCodePane = value;
                }
            }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            private string fileName;
            public bool FailFileName { get; set; }
            public string FileName { get { if (FailFileName) throw new InvalidOperationException("path unavailable"); return fileName; } set { fileName = value; } }
            public int Mode { get; set; }
            public FakeComponents VBComponents { get; } = new FakeComponents();
        }

        public sealed class FakeComponents : List<FakeComponent>
        {
            public FakeComponent Item(string name) { return this.Single(item => item.Name == name); }
        }

        public sealed class BrowserSnapshot
        {
            public BrowserWindowSnapshot[] Windows { get; set; }
        }

        public sealed class BrowserWindowSnapshot
        {
            public IDictionary<string, object> Properties { get; set; }
        }

        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
        }

        public sealed class FakeModule
        {
            public FakeComponent Parent { get; }
            public FakePane CodePane { get; }
            public FakeLines Lines { get; }
            public FakeProcedureBody ProcBodyLine { get; } = new FakeProcedureBody();
            public FakeProcedureLines ProcOfLine { get; } = new FakeProcedureLines();
            public string Code { get; set; }
            public int CountOfLines => string.IsNullOrEmpty(Code) ? 0 : Code.Split(new[] { "\r\n" }, StringSplitOptions.None).Length;

            public FakeModule(FakeComponent parent, string code)
            {
                Parent = parent;
                Code = code;
                CodePane = new FakePane
                {
                    CodeModule = this
                };
                Lines = new FakeLines(this);
            }
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
                    var lines = module.Code.Split(new[] { "\r\n" }, StringSplitOptions.None);
                    return string.Join("\r\n", lines.Skip(start - 1).Take(count));
                }
            }
        }

        public sealed class FakeProcedureBody
        {
            public int Body = 1;
            public int this[string procedure, int kind] => Body;
        }

        public sealed class FakeProcedureLines : System.Dynamic.DynamicObject
        {
            public Func<int, string> NameAtLine { get; set; } = line => "TryMe";
            public Func<int, int> KindAtLine { get; set; } = line => 0;
            public override bool TryGetIndex(System.Dynamic.GetIndexBinder binder, object[] indexes, out object result)
            {
                int line = Convert.ToInt32(indexes[0]);
                indexes[1] = KindAtLine(line);
                result = NameAtLine(line);
                return true;
            }
        }

        public sealed class FakeCodeWindow { public int FocusCount { get; private set; } public void SetFocus() { FocusCount++; } }
        public sealed class FakePane
        {
            public FakeCodeWindow Window { get; } = new FakeCodeWindow();
            public FakeModule CodeModule { get; set; }
            public int ShowCount { get; private set; }
            public bool FailGetSelection { get; set; }
            public bool RetainSelection { get; set; } = true;
            public Func<int[], int[]> SelectionReadback { get; set; }
            public Action OnShow { get; set; }
            public Action OnSetSelection { get; set; }
            public int StartLine { get; private set; } = 1;
            public int StartColumn { get; private set; } = 1;
            public int EndLine { get; private set; } = 1;
            public int EndColumn { get; private set; } = 1;

            public void Show()
            {
                ShowCount++;
                OnShow?.Invoke();
            }

            public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
            {
                OnSetSelection?.Invoke();
                if (!RetainSelection)
                {
                    StartLine = 1;
                    StartColumn = 1;
                    EndLine = 1;
                    EndColumn = 1;
                    return;
                }

                StartLine = startLine;
                StartColumn = startColumn;
                EndLine = endLine;
                EndColumn = endColumn;
            }

            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                if (FailGetSelection)
                    throw new InvalidOperationException("selection unavailable");
                startLine = StartLine;
                startColumn = StartColumn;
                endLine = EndLine;
                endColumn = EndColumn;
                if (SelectionReadback != null)
                {
                    var actual = SelectionReadback(new[] { startLine, startColumn, endLine, endColumn });
                    startLine = actual[0]; startColumn = actual[1]; endLine = actual[2]; endColumn = actual[3];
                }
            }
        }

        public sealed class FakeBar
        {
            private string name;
            public bool FailName { get; set; }
            public string Name { get { if (FailName) throw new InvalidOperationException("bar unavailable"); return name; } set { name = value; } }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
        }

        public sealed class FakeControl
        {
            private string caption;
            public bool FailCaption { get; set; }
            public string Caption { get { if (FailCaption) throw new InvalidOperationException("control unavailable"); return caption; } set { caption = value; } }
            public int Id { get; set; }
            private bool enabled = true;
            public Action OnEnabledRead { get; set; }
            public bool Enabled { get { OnEnabledRead?.Invoke(); return enabled; } set => enabled = value; }
            public int ExecuteCount { get; private set; }
            public Action OnExecute { get; set; }
            private readonly List<FakeControl> controls = new List<FakeControl>();
            public bool FailChildren { get; set; }
            public List<FakeControl> Controls { get { if (FailChildren) throw new InvalidOperationException("button has no children"); return controls; } }

            public void Execute()
            {
                ExecuteCount++;
                OnExecute?.Invoke();
            }
        }
    }
}
