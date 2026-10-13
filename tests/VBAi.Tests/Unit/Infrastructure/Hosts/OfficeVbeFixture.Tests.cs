using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks startup project selection without activating a host or emitting a native command.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureStartupBindingTests
    {
        private const string Source = @"E:\owned\Disposable.docm";

        [TestMethod]
        public void WordBindsCanonicalDocumentAmongNormalAndGeneratedTemplateProjects()
        {
            var fixture = Fixture("Word");
            fixture.BindStartupProject(new[]
            {
                Row("Normal", @"H:\personal\Normal", null),
                Row("Project", null, Source.ToUpperInvariant()),
                Row("TemplateProject", @"E:\owned\WordLifecycleTemplate.dotx", null)
            });
            Assert.AreEqual(Source, fixture.Project);
        }

        [TestMethod]
        public void CanonicalWordPathDisambiguatesEqualProjectNames()
        {
            var fixture = Fixture("Word");
            fixture.BindStartupProject(new[] { Row("Project", null, Source), Row("Project", null, @"E:\foreign.docm") });
            Assert.AreEqual(Source, fixture.Project);
        }

        [DataTestMethod]
        [DataRow("absent")]
        [DataRow("null")]
        [DataRow("empty")]
        [DataRow("foreign")]
        [DataRow("duplicate")]
        [DataRow("filename-only")]
        public void WordNeverFallsBackFromMissingOrAmbiguousCanonicalDocumentIdentity(string state)
        {
            var fixture = Fixture("Word");
            var row = Row("Project", Source, state == "foreign" ? @"E:\foreign.docm" : state == "empty" ? "" : null);
            if (state == "absent" || state == "filename-only") row.Remove("HostPath");
            var rows = state == "duplicate" ? new[] { Row("Project", null, Source), Row("Other", null, Source) } : new[] { row };
            Assert.ThrowsException<AssertFailedException>(() => fixture.BindStartupProject(rows));
            Assert.IsNull(fixture.Project);
        }

        [DataTestMethod]
        [DataRow("PowerPoint", true)]
        [DataRow("PowerPoint", false)]
        [DataRow("Access", true)]
        [DataRow("Access", false)]
        public void OtherHostsKeepExistingFileAndSoleProjectSelection(string host, bool path)
        {
            var fixture = Fixture(host);
            fixture.BindStartupProject(new[] { Row("Project", path ? Source : null, null) });
            Assert.AreEqual(path ? Source : "Project", fixture.Project);
        }

        private static OfficeVbeFixture Fixture(string host)
        {
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            foreach (var pair in new[] { new[] { "Kind", host }, new[] { "DocumentPath", Source } })
                typeof(OfficeVbeFixture).GetField("<" + pair[0] + ">k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, pair[1]);
            return fixture;
        }

        private static IDictionary<string, object> Row(string name, string file, string host)
        {
            return new Dictionary<string, object> { ["Name"] = name, ["FileName"] = file, ["HostPath"] = host };
        }
    }
}
