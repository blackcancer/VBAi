using System;
using System.Collections;
using System.Collections.Generic;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Public dynamic doubles that expose acquired getters and forbid native-style enumeration.</summary>
    public static class ComInspectionFixture
    {
        /// <summary>A one-based collection with observable, independently failing Count and Item getters.</summary>
        public sealed class Collection : IEnumerable
        {
            public readonly List<object> Values = new List<object>();
            public readonly List<int> ItemReads = new List<int>();
            public int CountReads, EnumeratorReads, FailingIndex;
            public int? ReportedCount;
            public bool FailCount;
            public Action BeforeCount, BeforeItem;
            public object Parent;

            public int Count
            {
                get
                {
                    CountReads++; BeforeCount?.Invoke();
                    if (FailCount) throw new InvalidOperationException("Count refused");
                    return ReportedCount ?? Values.Count;
                }
            }

            public object this[int index]
            {
                get
                {
                    ItemReads.Add(index); BeforeItem?.Invoke();
                    if (index == FailingIndex) throw new InvalidOperationException("Item refused");
                    if (index < 1 || index > Values.Count) throw new InvalidOperationException("Invalid one-based index");
                    return Values[index - 1];
                }
            }

            public IEnumerator GetEnumerator()
            {
                EnumeratorReads++;
                throw new InvalidOperationException("A native reader must not create a hidden enumerator");
            }
        }

        /// <summary>A graph of scalar properties and explicitly observable object-valued getters.</summary>
        public sealed class Node
        {
            public readonly Dictionary<string, object> Objects = new Dictionary<string, object>();
            public readonly List<string> ObjectReads = new List<string>();
            public readonly HashSet<string> FailingGetters = new HashSet<string>();
            public string Name = "Project", FileName = "C:\\Disposable.xlsm", Version = "7.1";
            public string Caption = "Module1", ProgId = "VBAi.AddIn", Guid = "test", Description = "VBAi";
            public int HWnd = 42, Type, WindowState, Left, Top, Width = 320, Height = 200;
            public int CodePaneView, TopLine = 1, CountOfVisibleLines = 20;
            public bool Visible = true, Connect = true;

            private object Get(string name)
            {
                ObjectReads.Add(name);
                if (FailingGetters.Contains(name)) throw new InvalidOperationException(name + " refused");
                return Objects.TryGetValue(name, out object value) ? value : null;
            }

            public object References => Get(nameof(References));
            public object Windows => Get(nameof(Windows));
            public object CodePanes => Get(nameof(CodePanes));
            public object VBProjects => Get(nameof(VBProjects));
            public object AddIns => Get(nameof(AddIns));
            public object ActiveWindow => Get(nameof(ActiveWindow));
            public object ActiveCodePane => Get(nameof(ActiveCodePane));
            public object ActiveVBProject => Get(nameof(ActiveVBProject));
            public object CodeModule => Get(nameof(CodeModule));
            public object Parent => Get(nameof(Parent));
            public object Collection => Get(nameof(Collection));
            public object Window => Get(nameof(Window));

            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                startLine = 1; startColumn = 2; endLine = 3; endColumn = 4;
            }
        }

        /// <summary>Reference metadata with optional getter failures and reentrant state-change callbacks.</summary>
        public sealed class Reference
        {
            public string LibraryName = "Library", Path = "C:\\Library.tlb", GuidText = "{00020430-0000-0000-C000-000000000046}";
            public int Major = 2, Minor;
            public bool Broken, BuiltIn, FailName, FailPath, FailGuid;
            public int NameReads, PathReads, BrokenReads, GuidReads;
            public Action BeforeBroken, BeforePath;

            public bool IsBroken { get { BrokenReads++; BeforeBroken?.Invoke(); return Broken; } }
            public string Name { get { NameReads++; if (FailName) throw new InvalidOperationException("Name refused"); return LibraryName; } }
            public string FullPath { get { PathReads++; BeforePath?.Invoke(); if (FailPath) throw new InvalidOperationException("Path refused"); return Path; } }
            // GUID is deliberately a property so a mandatory metadata read can fail after acquisition.
            public string GUID { get { GuidReads++; if (FailGuid) throw new InvalidOperationException("GUID refused"); return GuidText; } }
        }
    }
}
