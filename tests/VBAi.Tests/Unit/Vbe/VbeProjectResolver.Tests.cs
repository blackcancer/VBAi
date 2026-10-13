namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ProjectResolverTests
    {
        [TestMethod]
        public void AmbiguousProjectNamesRequireTheExactSavedDocumentPath()
        {
            var first = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\First.xlsm"
            };
            var second = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Second.xlsm"
            };
            var host = new FakeVbe
            {
                VBProjects = new List<FakeProject>
                {
                    first,
                    second
                }
            };
            Assert.ThrowsException<ArgumentException>(() => VbeProjectResolver.Resolve(host, " "));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(host, "VBAProject"));
            Assert.AreSame(second, (object)VbeProjectResolver.Resolve(host, second.FileName.ToLowerInvariant()));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(host, @"C:\Temp\Missing.xlsm"));
        }

        [TestMethod]
        public void NameLookupIsCaseInsensitiveForAUniqueProject()
        {
            var project = new FakeProject
            {
                Name = "Été",
                FileName = Path.Combine(Path.GetTempPath(), "été.xlsm")
            };
            var host = new FakeVbe
            {
                VBProjects = new List<FakeProject>
                {
                    project
                }
            };
            Assert.AreSame(project, (object)VbeProjectResolver.Resolve(host, "ÉTÉ"));
        }

        [TestMethod]
        public void SavedPathLookupSkipsUnsavedAndMalformedComFileNames()
        {
            var host = new HeterogeneousVbe();
            host.VBProjects.Add(new UnsavedProject());
            foreach (string name in new[] { null, " ", "invalid\0file", "bad:filename", @"C:\" + new string('x', 40000) })
                host.VBProjects.Add(new FakeProject { Name = "Other", FileName = name });
            var selected = new FakeProject { Name = "Selected", FileName = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "disposable-selector.xlsm") };
            host.VBProjects.Add(selected);
            Assert.AreSame(selected, (object)VbeProjectResolver.Resolve(host, selected.FileName));
        }

        /// <summary>Resolving one project cannot retain unrelated coverage copies through enumeration getters.</summary>
        [DataTestMethod]
        [DataRow("name")][DataRow("path")][DataRow("missing")][DataRow("ambiguous")]
        [DataRow("duplicate")][DataRow("unreadable")][DataRow("enumeration")]
        public void ResolverReleasesEveryNonTransferredAcquisition(string scenario)
        {
            var first = new FakeProject { Name = "Selected", FileName = @"C:\owned\Original.xlsm" };
            var copy = new FakeProject { Name = scenario == "ambiguous" ? "Selected" : "Copy", FileName = @"C:\owned\coverage.xlsm" };
            var items = new ResolverCollection(first, scenario == "duplicate" ? first : (object)copy);
            items.ThrowAfterFirst = scenario == "enumeration";
            var host = new ResolverHost { VBProjects = items };
            var released = new List<object>();
            Func<object, string> path = value =>
            {
                if (scenario == "unreadable" && ReferenceEquals(value, copy)) throw new IOException("Closed temporary project.");
                return ((FakeProject)value).FileName;
            };
            string selector = scenario == "missing" ? "Absent" : scenario == "path" || scenario == "unreadable" ? first.FileName : first.Name;
            bool failure = scenario == "missing" || scenario == "ambiguous" || scenario == "duplicate" || scenario == "enumeration";
            if (failure)
                Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(host, selector, path, released.Add));
            else
                Assert.AreSame(first, (object)VbeProjectResolver.Resolve(host, selector, path, released.Add));

            var expected = new List<object>();
            if (!items.ThrowAfterFirst) expected.Add(scenario == "duplicate" ? first : (object)copy);
            if (failure) expected.Add(first);
            expected.Add(items.Enumerator); expected.Add(items);
            CollectionAssert.AreEqual(expected, released, "Only the unique returned acquisition transfers; even refusals release in reverse order.");
            Assert.IsFalse(released.Contains(host), "The borrowed VBE is never released.");
        }

        public sealed class ResolverHost
        {
            public ResolverCollection VBProjects { get; set; }
        }

        public sealed class ResolverCollection : System.Collections.IEnumerable
        {
            internal readonly object[] Values;
            internal bool ThrowAfterFirst;
            internal ResolverEnumerator Enumerator;
            internal ResolverCollection(params object[] values) { Values = values; }
            public System.Collections.IEnumerator GetEnumerator() => Enumerator = new ResolverEnumerator(this);
        }

        public sealed class ResolverEnumerator : System.Collections.IEnumerator
        {
            private readonly ResolverCollection source;
            private int position = -1;
            internal ResolverEnumerator(ResolverCollection source) { this.source = source; }
            public object Current => source.Values[position];
            public bool MoveNext()
            {
                if (source.ThrowAfterFirst && position == 0) throw new InvalidOperationException("Collection changed.");
                return ++position < source.Values.Length;
            }
            public void Reset() { position = -1; }
        }
    }
}
