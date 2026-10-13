using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks shutdown toolbar acquisition ownership without Office or native mutation.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ToolbarShutdownOwnershipTests
    {
        /// <summary>Checks getter, deletion, iterator and release boundaries as one complete matrix.</summary>
        /// <param name="variation">Boundary selected for this isolated managed fixture.</param>
        [DataTestMethod]
        [DataRow("native indexed")]
        [DataRow("managed enumeration")]
        [DataRow("empty bars")]
        [DataRow("menu skipped")]
        [DataRow("empty controls")]
        [DataRow("reverse owned deletion")]
        [DataRow("mixed commands")]
        [DataRow("same control acquired twice")]
        [DataRow("null tag")]
        [DataRow("empty tag")]
        [DataRow("uppercase tag")]
        [DataRow("short tag")]
        [DataRow("long tag")]
        [DataRow("trailing newline tag")]
        [DataRow("legacy tag")]
        [DataRow("bars getter failure")]
        [DataRow("bar item failure")]
        [DataRow("bar type failure")]
        [DataRow("controls getter failure")]
        [DataRow("controls count failure")]
        [DataRow("control item failure")]
        [DataRow("tag getter failure")]
        [DataRow("delete failure")]
        [DataRow("release failure")]
        public void ShutdownCleanupBalancesOnlyAcquiredReferencesAndNeverRetriesAnUncertainCall(string variation)
        {
            var state = new State(variation);
            var host = new Host(state);
            var service = new VbeEditorWindows(host)
            {
                IndexedCollection = value => variation != "managed enumeration",
                ReleaseAcquired = state.Release
            };
            Exception error = null;
            try { service.RemoveTemporaryToolbarCommands(); }
            catch (Exception failure) { error = failure; }
            bool fails = variation.EndsWith("failure", StringComparison.Ordinal);
            if (fails) Assert.AreSame(state.Failure, error, "The original injected failure must be preserved.");
            else Assert.IsNull(error);
            Assert.IsFalse(state.Released.Any(value => ReferenceEquals(value, host)), "The editor is borrowed.");
            Assert.AreEqual(1, host.BarsReads, "CommandBars must be acquired exactly once.");
            if (variation == "release failure")
            {
                Assert.AreEqual(1, state.ReleaseAttempts);
                Assert.AreEqual(0, state.Released.Count);
            }
            else
            {
                Assert.AreEqual(state.Acquired.Count, state.Released.Count, "Partial acquisitions must also be balanced.");
                for (int i = 0; i < state.Acquired.Count; ++i)
                    Assert.AreSame(state.Acquired[state.Acquired.Count - 1 - i], state.Released[i]);
            }
            Assert.AreEqual(variation == "managed enumeration" ? 1 : 0, state.IteratorCreates,
                "Native indexed access must not create an implicit COM enumerator.");
            Assert.AreEqual(state.IteratorCreates, state.IteratorDisposals);
            foreach (var bar in host.Bars.Values)
                Assert.IsTrue(bar.ControlsReads <= 1, "Acquire Controls once, not once per index.");
            if (variation == "reverse owned deletion") CollectionAssert.AreEqual(new[] { 3, 2, 1 }, state.Deleted.ToArray());
            if (variation == "delete failure") Assert.AreEqual(1, state.DeleteAttempts, "Deletion must never be replayed.");
            if (!fails && variation != "reverse owned deletion")
            {
                int expected = variation == "native indexed" || variation == "managed enumeration" || variation == "mixed commands" ? 1 : 0;
                Assert.AreEqual(expected, state.Deleted.Count);
            }
            if (variation == "mixed commands")
            {
                Assert.AreEqual(1, host.Bars.Values[0].ControlsValue.Values.Count);
                Assert.AreEqual("Foreign.Command", host.Bars.Values[0].ControlsValue.Values[0].Tag);
            }
            if (variation == "same control acquired twice")
            {
                object same = host.Bars.Values[0].ControlsValue.Values[0];
                Assert.AreEqual(2, state.Released.Count(value => ReferenceEquals(value, same)),
                    "Release per getter acquisition, not per distinct RCW.");
            }
        }

        public sealed class State
        {
            internal readonly string Variation;
            internal readonly int Owner = Thread.CurrentThread.ManagedThreadId;
            internal readonly Exception Failure = new InvalidOperationException("Original shutdown fixture failure.");
            internal readonly List<object> Acquired = new List<object>(), Released = new List<object>();
            internal readonly List<int> Deleted = new List<int>();
            internal int ReleaseAttempts, DeleteAttempts, IteratorCreates, IteratorDisposals;
            internal State(string variation) { Variation = variation; }
            internal object Acquire(object value) { Acquired.Add(value); return value; }
            internal void Release(object value)
            {
                Assert.AreEqual(Owner, Thread.CurrentThread.ManagedThreadId);
                ReleaseAttempts++;
                if (Variation == "release failure") throw Failure;
                Released.Add(value);
                if (value is IEnumerator && value is IDisposable iterator) iterator.Dispose();
            }
        }

        public sealed class Host
        {
            private readonly State state;
            internal readonly Bars Bars;
            internal int BarsReads;
            internal Host(State state) { this.state = state; Bars = new Bars(state); }
            public Bars CommandBars
            {
                get { BarsReads++; if (state.Variation == "bars getter failure") throw state.Failure; return (Bars)state.Acquire(Bars); }
            }
        }

        public sealed class Bars : IEnumerable
        {
            private readonly State state;
            internal readonly List<Bar> Values = new List<Bar>();
            internal Bars(State state) { this.state = state; if (state.Variation != "empty bars") Values.Add(new Bar(state)); }
            public int Count { get { return Values.Count; } }
            public Bar this[int index]
            {
                get { if (state.Variation == "bar item failure") throw state.Failure; return (Bar)state.Acquire(Values[index - 1]); }
            }
            public IEnumerator GetEnumerator()
            {
                state.IteratorCreates++;
                var iterator = new Iterator(state, Values.Cast<object>().ToArray());
                state.Acquire(iterator);
                return iterator;
            }
        }

        public sealed class Iterator : IEnumerator, IDisposable
        {
            private readonly State state;
            private readonly object[] items;
            private int index = -1;
            internal Iterator(State state, object[] items) { this.state = state; this.items = items; }
            public bool MoveNext() { return ++index < items.Length; }
            public object Current { get { return state.Acquire(items[index]); } }
            public void Reset() { throw new NotSupportedException(); }
            public void Dispose() { state.IteratorDisposals++; }
        }

        public sealed class Bar
        {
            private readonly State state;
            internal readonly Controls ControlsValue;
            internal int ControlsReads;
            internal Bar(State state) { this.state = state; ControlsValue = new Controls(state); }
            public int Type
            {
                get { if (state.Variation == "bar type failure") throw state.Failure; return state.Variation == "menu skipped" ? 1 : 0; }
            }
            public Controls Controls
            {
                get { ControlsReads++; if (state.Variation == "controls getter failure") throw state.Failure; return (Controls)state.Acquire(ControlsValue); }
            }
        }

        public sealed class Controls
        {
            private readonly State state;
            internal readonly List<Button> Values = new List<Button>();
            internal Controls(State state)
            {
                this.state = state;
                if (state.Variation == "empty controls") return;
                string tag = "Foreign.Command";
                if (state.Variation == "native indexed" || state.Variation == "managed enumeration" ||
                    state.Variation == "reverse owned deletion" || state.Variation == "delete failure" || state.Variation == "mixed commands")
                    tag = "VBAi.ToolbarCommand." + new string('a', 32);
                if (state.Variation == "null tag") tag = null;
                if (state.Variation == "empty tag") tag = "";
                if (state.Variation == "uppercase tag") tag = "VBAi.ToolbarCommand." + new string('A', 32);
                if (state.Variation == "short tag") tag = "VBAi.ToolbarCommand." + new string('a', 31);
                if (state.Variation == "long tag") tag = "VBAi.ToolbarCommand." + new string('a', 33);
                if (state.Variation == "trailing newline tag") tag = "VBAi.ToolbarCommand." + new string('a', 32) + "\n";
                if (state.Variation == "legacy tag") tag = "CodexVBE.ToolbarCommand." + new string('a', 32);
                Values.Add(new Button(state, this, tag, 1));
                if (state.Variation == "reverse owned deletion")
                { Values.Add(new Button(state, this, tag, 2)); Values.Add(new Button(state, this, tag, 3)); }
                if (state.Variation == "same control acquired twice") Values.Add(Values[0]);
                if (state.Variation == "mixed commands") Values.Add(new Button(state, this, "Foreign.Command", 2));
            }
            public int Count { get { if (state.Variation == "controls count failure") throw state.Failure; return Values.Count; } }
            public Button this[int index]
            {
                get { if (state.Variation == "control item failure") throw state.Failure; return (Button)state.Acquire(Values[index - 1]); }
            }
        }

        public sealed class Button
        {
            private readonly State state;
            private readonly Controls owner;
            private readonly string tag;
            private readonly int originalIndex;
            internal Button(State state, Controls owner, string tag, int originalIndex)
            { this.state = state; this.owner = owner; this.tag = tag; this.originalIndex = originalIndex; }
            public string Tag { get { if (state.Variation == "tag getter failure") throw state.Failure; return tag; } }
            public void Delete()
            {
                state.DeleteAttempts++;
                if (state.Variation == "delete failure") throw state.Failure;
                state.Deleted.Add(originalIndex);
                owner.Values.Remove(this);
            }
        }
    }
}
