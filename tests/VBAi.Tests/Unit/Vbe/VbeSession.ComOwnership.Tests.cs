using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks that native module reads leave no getter acquisition owned by a finished request.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeSessionModuleOwnershipTests
    {
        public sealed class CodeLines
        {
            public string Text;
            public bool Fault;
            public int Reads;
            public string this[int start, int count]
            {
                get
                {
                    Assert.AreEqual(1, start); Assert.AreEqual(2, count);
                    Reads++;
                    if (Fault) throw new InvalidOperationException("Native source unavailable.");
                    return Text;
                }
            }
        }

        public sealed class CodeModule
        {
            public string Fault;
            public CodeLines Lines { get; set; }
            public int CountOfLines => Fault == "CountOfLines" ? throw new InvalidOperationException("Native lines unavailable.") :
                string.IsNullOrEmpty(Lines.Text) ? 0 : 2;
        }

        public sealed class Component
        {
            public string Fault, ComponentName;
            public CodeModule Source;
            public string Name => Fault == "Name" ? throw new InvalidOperationException("Native name unavailable.") : ComponentName;
            public int Type => Fault == "Type" ? throw new InvalidOperationException("Native type unavailable.") : 1;
            public object CodeModule => Fault == "CodeModule" ? throw new InvalidOperationException("Native module unavailable.") : Source;
        }

        public sealed class Project
        {
            public bool Fault;
            public ProjectResolverTests.ResolverCollection Components;
            public object VBComponents => Fault ? throw new InvalidOperationException("Native components unavailable.") : Components;
        }

        /// <summary>Success, empty code, missing targets and getter failures all retain the original metadata/read contract and balance acquisitions.</summary>
        [DataTestMethod]
        [DataRow("success")][DataRow("empty")][DataRow("missing")][DataRow("enumeration")]
        [DataRow("VBComponents")][DataRow("Name")][DataRow("CodeModule")][DataRow("CountOfLines")][DataRow("Lines")]
        public void ReadModuleBalancesAllAcquisitionsWithoutReleasingAnotherOwner(string scenario)
        {
            var source = new CodeModule { Lines = new CodeLines { Text = scenario == "empty" ? "" : "Option Explicit\r\n' Été", Fault = scenario == "Lines" }, Fault = scenario };
            var other = new Component { ComponentName = "Other", Source = source };
            var target = new Component { ComponentName = scenario == "missing" ? "Different" : "ModuleÉté", Source = source, Fault = scenario };
            var collection = new ProjectResolverTests.ResolverCollection(other, target) { ThrowAfterFirst = scenario == "enumeration" };
            var project = new Project { Components = collection, Fault = scenario == "VBComponents" };
            var released = new List<object>(); int resolutions = 0;
            Func<object> resolve = () => { resolutions++; return project; };
            if (scenario == "success" || scenario == "empty")
            {
                dynamic result = VbeSession.ReadModuleSnapshot(resolve, @"C:\owned\Original.docm", "moduleété", released.Add);
                Assert.AreEqual(@"C:\owned\Original.docm", (string)result.Project);
                Assert.AreEqual("moduleété", (string)result.Module);
                Assert.AreEqual(source.Lines.Text, (string)result.Code);
                Assert.AreEqual(EditorDocument.Hash(source.Lines.Text), (string)result.Sha256);
                Assert.AreEqual(scenario == "empty" ? 0 : 1, source.Lines.Reads);
            }
            else
                Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadModuleSnapshot(resolve, "Owned", "moduleété", released.Add));
            var expected = new List<object>();
            if (scenario == "success" || scenario == "empty" || scenario == "CountOfLines" || scenario == "Lines") expected.Add(source);
            if (scenario != "VBComponents" && scenario != "enumeration") expected.Add(target);
            if (scenario != "VBComponents") { expected.Add(other); expected.Add(collection.Enumerator); expected.Add(collection); }
            expected.Add(project);
            CollectionAssert.AreEqual(expected, released, "Only this request's acquisitions are released, once each and in reverse order.");
            Assert.AreEqual(1, resolutions, "A failed read never reacquires the project.");
        }

        /// <summary>Inventory releases even duplicate code-module acquisitions independently and preserves getter errors.</summary>
        [DataTestMethod]
        [DataRow("success")][DataRow("shared")][DataRow("enumeration")][DataRow("VBComponents")]
        [DataRow("Name")][DataRow("Type")][DataRow("CodeModule")][DataRow("CountOfLines")]
        public void ComponentInventoryBalancesEveryGetterIncludingRepeatedRcw(string scenario)
        {
            var firstSource = new CodeModule { Lines = new CodeLines { Text = "Option Explicit\r\n' First" } };
            var lastSource = scenario == "shared" ? firstSource : new CodeModule { Lines = new CodeLines { Text = "Option Explicit\r\n' Last" }, Fault = scenario };
            var first = new Component { ComponentName = "First", Source = firstSource };
            var last = new Component { ComponentName = "Last", Source = lastSource, Fault = scenario };
            var collection = new ProjectResolverTests.ResolverCollection(first, last) { ThrowAfterFirst = scenario == "enumeration" };
            var project = new Project { Components = collection, Fault = scenario == "VBComponents" };
            var released = new List<object>(); int resolutions = 0;
            Func<object> resolve = () => { resolutions++; return project; };
            if (scenario == "success" || scenario == "shared")
            {
                var results = (List<object>)VbeSession.ReadComponentInventory(resolve, released.Add);
                Assert.AreEqual(2, results.Count);
                dynamic firstRow = results[0], lastRow = results[1];
                Assert.AreEqual("First", (string)firstRow.Name); Assert.AreEqual("Last", (string)lastRow.Name);
                Assert.AreEqual(1, (int)firstRow.Type); Assert.AreEqual(2, (int)lastRow.Lines);
            }
            else Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadComponentInventory(resolve, released.Add));
            var expected = new List<object>();
            if (scenario == "success" || scenario == "shared" || scenario == "CountOfLines") expected.Add(lastSource);
            if (scenario != "enumeration" && scenario != "VBComponents") expected.Add(last);
            if (scenario != "VBComponents") { expected.Add(firstSource); expected.Add(first); expected.Add(collection.Enumerator); expected.Add(collection); }
            expected.Add(project);
            CollectionAssert.AreEqual(expected, released);
            Assert.AreEqual(scenario == "shared" ? 2 : scenario == "VBComponents" ? 0 : 1, released.Count(item => ReferenceEquals(item, firstSource)));
            Assert.AreEqual(1, resolutions);
            Assert.AreEqual(0, firstSource.Lines.Reads, "Inventory reads line counts, not source text.");
        }

        /// <summary>Invalid requests and resolution failure acquire nothing; a release fault is preserved without retrying it.</summary>
        [TestMethod]
        public void ReadGuardsAndReleaseFailureNeverRetryTheNativeBoundary()
        {
            int resolutions = 0;
            Func<object> forbidden = () => { resolutions++; throw new InvalidOperationException("Resolution failed."); };
            Assert.ThrowsException<ArgumentException>(() => VbeSession.ReadModuleSnapshot(forbidden, "P", " "));
            Assert.ThrowsException<ArgumentNullException>(() => VbeSession.ReadModuleSnapshot(null, "P", "M"));
            Assert.ThrowsException<ArgumentNullException>(() => VbeSession.ReadComponentInventory(null));
            Assert.AreEqual(0, resolutions);
            Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadModuleSnapshot(forbidden, "P", "M"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadComponentInventory(forbidden));
            Assert.AreEqual(2, resolutions);
            var source = new CodeModule { Lines = new CodeLines { Text = "Option Explicit\r\n' Test" } };
            var component = new Component { ComponentName = "M", Source = source };
            var project = new Project { Components = new ProjectResolverTests.ResolverCollection(component) };
            int releases = 0;
            var error = Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadModuleSnapshot(() => project, "P", "M", item =>
            {
                Assert.AreSame(source, item); releases++; throw new InvalidOperationException("Release uncertain.");
            }));
            Assert.AreEqual("Release uncertain.", error.Message); Assert.AreEqual(1, releases);
        }
    }
}
