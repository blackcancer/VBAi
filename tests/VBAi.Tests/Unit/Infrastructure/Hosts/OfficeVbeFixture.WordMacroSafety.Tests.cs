using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureWordMacroSafetyTests
    {
        [TestMethod]
        public void SuppressionPrecedesDocumentCreationAndReleasesItsExactLeaseOnce()
        {
            var app = new FakeApplication();
            OfficeVbeFixture.SuppressWordAutoMacros(app, () => app.Trace.Add("owner"), lease => {
                Assert.AreSame(app.Commands, lease); app.Trace.Add("release");
            }, unused => Assert.Fail("A successful command must not be retained as uncertain."));
            app.Trace.Add("create");
            CollectionAssert.AreEqual(new[] { "owner", "security:3", "wordbasic", "disable", "release", "create" }, app.Trace);
            Assert.AreEqual(1, app.Commands.Calls);
        }

        [DataTestMethod]
        [DataRow("owner")]
        [DataRow("security")]
        [DataRow("wordbasic")]
        [DataRow("disable")]
        [DataRow("release")]
        public void FailurePreventsDocumentWorkAndNeverReplaysOrReenables(string phase)
        {
            var app = new FakeApplication { FailurePhase = phase };
            var error = app.Error;
            var retained = new List<object>();
            var observed = Assert.ThrowsException<InvalidOperationException>(() => {
                OfficeVbeFixture.SuppressWordAutoMacros(app, () => {
                    app.Trace.Add("owner"); if (phase == "owner") throw error;
                }, lease => {
                    Assert.AreSame(app.Commands, lease); app.Trace.Add("release");
                    if (phase == "release") throw error;
                }, retained.Add);
                app.Trace.Add("create");
            });
            Assert.AreSame(error, observed);
            Assert.IsFalse(app.Trace.Contains("create"));
            Assert.AreEqual(phase == "disable" || phase == "release" ? 1 : 0, app.Commands.Calls);
            Assert.AreEqual(phase == "disable" || phase == "release" ? 1 : 0, retained.Count);
            if (retained.Count != 0) Assert.AreSame(app.Commands, retained[0]);
            int failingIndex = Array.IndexOf(new[] { "owner", "security:3", "wordbasic", "disable", "release" },
                phase == "security" ? "security:3" : phase);
            Assert.AreEqual(failingIndex + 1, app.Trace.Count, "No command may follow the uncertain phase.");
        }

        [TestMethod]
        public void MissingWordBasicRefusesCreationWithoutDisablingOrCleanup()
        {
            var app = new FakeApplication { MissingWordBasic = true };
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.SuppressWordAutoMacros(app,
                () => app.Trace.Add("owner"), unused => Assert.Fail("No lease exists."), unused => Assert.Fail("No lease exists.")));
            CollectionAssert.AreEqual(new[] { "owner", "security:3", "wordbasic" }, app.Trace);
            Assert.AreEqual(0, app.Commands.Calls);
        }

        [TestMethod]
        public void MissingDependenciesRefuseBeforeAnyApplicationAccess()
        {
            var app = new FakeApplication(); Action owner = () => Assert.Fail("No ownership work after invalid arguments.");
            Action<object> cleanup = unused => Assert.Fail("No cleanup after invalid arguments.");
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.SuppressWordAutoMacros(null, owner, cleanup, cleanup));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.SuppressWordAutoMacros(app, null, cleanup, cleanup));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.SuppressWordAutoMacros(app, owner, null, cleanup));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.SuppressWordAutoMacros(app, owner, cleanup, null));
            Assert.AreEqual(0, app.Trace.Count);
        }

        public sealed class FakeApplication
        {
            public readonly List<string> Trace = new List<string>();
            public readonly InvalidOperationException Error = new InvalidOperationException("Synthetic command uncertainty.");
            public readonly FakeWordBasic Commands;
            public string FailurePhase;
            public bool MissingWordBasic;
            public FakeApplication() { Commands = new FakeWordBasic(this); }
            public int AutomationSecurity { set { Trace.Add("security:" + value); if (FailurePhase == "security") throw Error; Assert.AreEqual(3, value); } }
            public object WordBasic { get { Trace.Add("wordbasic"); if (FailurePhase == "wordbasic") throw Error; return MissingWordBasic ? null : Commands; } }
        }

        public sealed class FakeWordBasic
        {
            private readonly FakeApplication app;
            public int Calls;
            public FakeWordBasic(FakeApplication app) { this.app = app; }
            public void DisableAutoMacros() { Calls++; app.Trace.Add("disable"); if (app.FailurePhase == "disable") throw app.Error; }
        }
    }
}
