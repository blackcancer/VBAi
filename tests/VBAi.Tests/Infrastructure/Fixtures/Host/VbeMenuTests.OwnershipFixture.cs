using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests
{
    /// <summary>Public dynamic graph with counted one-based getters and no hidden enumerator.</summary>
    public sealed class MenuOwnershipGraph
    {
        public readonly List<object> Acquired = new List<object>(), Released = new List<object>();
        public readonly List<Button> Buttons = new List<Button>();
        public readonly List<Delegate> Callbacks = new List<Delegate>();
        public readonly Collection Bars;
        public readonly Bar MenuBar;
        public readonly Popup View, Tools, Help;
        public int FailingButton, FailingSubscription, SubscriptionEntries, CallbackEntries;
        public string FailingField;
        public bool FailRelease, FailDelete;

        public MenuOwnershipGraph()
        {
            Bars = new Collection(this);
            MenuBar = new Bar(this) { TypeValue = 1, Name = "Menu Bar" };
            View = new Popup(this, "&View"); Tools = new Popup(this, "&Tools"); Help = new Popup(this, "?");
            MenuBar.Children.Values.AddRange(new object[] { View, Tools, Help });
            Bars.Values.Add(MenuBar);
            Bars.Values.Add(new Bar(this) { TypeValue = 2, Name = "Code Window" });
        }

        public object CommandBars => Take(Bars);
        public object Take(object value) { Acquired.Add(value); return value; }
        public void Release(object value)
        {
            Released.Add(value);
            if (FailRelease) throw new InvalidOperationException("release refused");
        }
        internal VbaTestComReferences Scope() => new VbaTestComReferences(Release, value => value is Collection);
        public void VerifyBalanced()
        {
            Assert.AreEqual(Acquired.Count, Released.Count);
            for (int index = 0; index < Acquired.Count; index++)
                Assert.AreSame(Acquired[Acquired.Count - index - 1], Released[index], "One release per successful getter, including aliases.");
            foreach (var collection in Acquired.OfType<Collection>().Distinct())
                Assert.AreEqual(0, collection.EnumeratorReads);
        }

        internal VbeMenu Create(bool optional, List<object> subscribed = null, List<object> removed = null)
        {
            Action callback = () => CallbackEntries++;
            return new VbeMenu(this, callback, callback, callback, optional ? new Action<string>(_ => CallbackEntries++) : null,
                (button, iid, dispid, handler) =>
                {
                    SubscriptionEntries++;
                    if (SubscriptionEntries == FailingSubscription) throw new InvalidOperationException("subscribe refused");
                    subscribed?.Add(button); Callbacks.Add(handler);
                },
                (button, iid, dispid, handler) => removed?.Add(button), (button, type) => { },
                optional ? callback : null, optional ? callback : null, optional ? callback : null, callback,
                Release, value => value is Collection);
        }

        public sealed class Collection : IEnumerable
        {
            private readonly MenuOwnershipGraph graph;
            public readonly List<object> Values = new List<object>();
            public bool FailCount;
            public int? ReportedCount;
            public int FailingIndex, EnumeratorReads;
            public Collection(MenuOwnershipGraph graph) { this.graph = graph; }
            public int Count { get { if (FailCount) throw new InvalidOperationException("count refused"); return ReportedCount ?? Values.Count; } }
            public object this[int index]
            {
                get
                {
                    if (index == FailingIndex) throw new InvalidOperationException("item refused");
                    return graph.Take(Values[index - 1]);
                }
            }
            public object Add(int type, object id, object parameter, object before, bool temporary)
            {
                var button = new Button(graph);
                graph.Buttons.Add(button); Values.Add(button);
                return graph.Take(button);
            }
            public IEnumerator GetEnumerator() { EnumeratorReads++; throw new InvalidOperationException("hidden enumerator refused"); }
        }

        public sealed class Bar
        {
            private readonly MenuOwnershipGraph graph;
            public readonly Collection Children;
            public bool FailType;
            public int TypeValue;
            public string Name;
            public Bar(MenuOwnershipGraph graph) { this.graph = graph; Children = new Collection(graph); }
            public int Type { get { if (FailType) throw new InvalidOperationException("type refused"); return TypeValue; } }
            public object Controls => graph.Take(Children);
        }

        public sealed class Popup
        {
            private readonly MenuOwnershipGraph graph;
            public readonly Collection Children;
            public bool FailCaption;
            private string caption;
            public Popup(MenuOwnershipGraph graph, string caption) { this.graph = graph; this.caption = caption; Children = new Collection(graph); }
            public string Caption { get { if (FailCaption) throw new InvalidOperationException("caption refused"); return caption; } set { caption = value; } }
            public object Controls => graph.Take(Children);
        }

        public sealed class Button
        {
            private readonly MenuOwnershipGraph graph;
            private string caption, tag, tooltip;
            public int DeleteCount;
            public Button(MenuOwnershipGraph graph) { this.graph = graph; }
            private void Set(string field) { if (graph.FailingButton == graph.Buttons.IndexOf(this) + 1 && graph.FailingField == field) throw new InvalidOperationException("metadata refused"); }
            public string Caption { get => caption; set { Set("caption"); caption = value; } }
            public string Tag { get => tag; set { Set("tag"); tag = value; } }
            public string TooltipText { get => tooltip; set { Set("tooltip"); tooltip = value; } }
            public void Delete() { DeleteCount++; if (graph.FailDelete) throw new InvalidOperationException("delete refused"); }
        }
    }
}
