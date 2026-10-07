using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual selection and unsaved-workbook transition with managed launch and workbook fakes.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ExcelVbeFixtureBootstrapSelectionTests
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        private static void Set(ExcelVbeFixture fixture, string name, object value) => typeof(ExcelVbeFixture).GetField(name, Private).SetValue(fixture, value);
        private static object Get(ExcelVbeFixture fixture, string name) => typeof(ExcelVbeFixture).GetField(name, Private).GetValue(fixture);
        private static Exception Capture(Action action) { try { action(); return null; } catch (Exception error) { return error; } }

        private sealed class Scope : IDisposable
        {
            internal readonly string Root = Path.Combine(Path.GetTempPath(), "VBAi-BootstrapSelection-" + Guid.NewGuid().ToString("N"));
            internal readonly ExcelVbeFixture Fixture = (ExcelVbeFixture)Activator.CreateInstance(typeof(ExcelVbeFixture), true);
            internal readonly List<string> Calls = new List<string>();
            internal readonly Process Descriptor = Process.GetCurrentProcess();
            internal readonly IOException Primary = new IOException("synthetic uncertain workbook operation");
            internal readonly Workbook Seed;
            internal readonly Workbook Added;
            internal readonly Workbooks Books;
            internal readonly Application App;
            internal string Fault;
            internal Scope()
            {
                Directory.CreateDirectory(Root);
                typeof(ExcelVbeFixture).GetProperty("Root", Private | BindingFlags.Public).SetValue(Fixture, Root);
                typeof(ExcelVbeFixture).GetProperty("ProcessId", Private | BindingFlags.Public).SetValue(Fixture, Descriptor.Id);
                Seed = new Workbook(Calls, "seed", () => Fault, Primary);
                Added = new Workbook(Calls, "added", () => Fault, Primary);
                Books = new Workbooks(Calls, Added, () => Fault, Primary);
                App = new Application(Calls);
                Set(Fixture, "owned", true); Set(Fixture, "ownedProcess", Descriptor);
                Set(Fixture, "workbook", Seed); Set(Fixture, "workbooks", Books); Set(Fixture, "application", App);
            }
            internal string Trace => Path.Combine(Root, "inspection.jsonl");
            internal ExcelVbeFixture Start(string desktop, string setting, string trace = null)
                => ExcelVbeFixture.StartSelectedBootstrap(desktop, setting, trace ?? Trace,
                    actual => { Assert.AreEqual(desktop, actual); Calls.Add("desktop"); },
                    actual => { Assert.AreEqual(Trace, actual); Calls.Add("owned"); return Fixture; },
                    () => { Calls.Add("com"); return Fixture; });
            public void Dispose()
            {
                var retained = (IList)typeof(ExcelVbeFixture).GetField("retainedBootstraps", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                lock (retained) retained.Remove(Fixture);
                Descriptor.Dispose();
                Directory.Delete(Root, true);
            }
        }

        public sealed class Workbook
        {
            private readonly List<string> calls;
            private readonly string kind;
            private readonly Func<string> fault;
            private readonly Exception primary;
            public bool Closed;
            public int CloseCalls;
            public Workbook(List<string> calls, string kind, Func<string> fault, Exception primary)
            { this.calls = calls; this.kind = kind; this.fault = fault; this.primary = primary; }
            public void Close(bool save)
            {
                calls.Add(kind + ".Close:" + save); CloseCalls++;
                Assert.IsFalse(save);
                if (fault() == "close-before") throw primary;
                Closed = true;
                if (fault() == "close-after") throw primary;
            }
            public string Path { get { calls.Add(kind + ".Path"); if (fault() == "path-error") throw primary; return fault() == "saved-path" ? @"C:\AlreadySaved" : ""; } }
            public string Name { get { calls.Add(kind + ".Name"); if (fault() == "name-error") throw primary; return "Book1"; } }
        }

        public sealed class Workbooks
        {
            private readonly List<string> calls;
            private readonly Workbook added;
            private readonly Func<string> fault;
            private readonly Exception primary;
            public int AddCalls;
            public bool Added;
            public Workbooks(List<string> calls, Workbook added, Func<string> fault, Exception primary)
            { this.calls = calls; this.added = added; this.fault = fault; this.primary = primary; }
            public object Add()
            {
                calls.Add("Add"); AddCalls++;
                if (fault() == "add-before") throw primary;
                Added = true;
                if (fault() == "add-after") throw primary;
                return added;
            }
            public int Count { get { calls.Add("Count"); if (fault() == "count-error") throw primary; return fault() == "extra-workbook" ? 2 : 1; } }
        }

        public sealed class Application
        {
            private readonly List<string> calls;
            public Application(List<string> calls) { this.calls = calls; }
            public void Quit() { calls.Add("Quit"); }
        }

        [DataTestMethod]
        [DataRow(null, null, false)]
        [DataRow(null, "", false)]
        [DataRow(null, " ", false)]
        [DataRow(null, "0", false)]
        [DataRow(null, "true", false)]
        [DataRow(null, " 1 ", false)]
        [DataRow(null, "1 ", false)]
        [DataRow(null, "1", true)]
        [DataRow("", "1", true)]
        [DataRow(" ", "1", true)]
        [DataRow("Q027Private", null, true)]
        [DataRow("Q027Private", "", true)]
        [DataRow("Q027Private", "0", true)]
        [DataRow("Q027Private", "true", true)]
        [DataRow("Q027Private", " 1 ", true)]
        [DataRow("Q027Private", "1", true)]
        public void SelectionExecutesOnlyTheRequestedLauncherAndPreservesTheUnsavedPrecondition(string desktop, string setting, bool explicitLaunch)
        {
            using (var scope = new Scope())
            {
                Assert.AreSame(scope.Fixture, scope.Start(desktop, setting));
                bool privateDesktop = !string.IsNullOrWhiteSpace(desktop);
                var expected = new List<string>();
                if (privateDesktop) expected.Add("desktop");
                if (explicitLaunch) expected.AddRange(new[] { "owned", "seed.Close:False", "Add", "Count", "added.Path", "added.Name", "added.Path" });
                else expected.Add("com");
                CollectionAssert.AreEqual(expected, scope.Calls);
                Assert.AreEqual(explicitLaunch ? 1 : 0, scope.Seed.CloseCalls);
                Assert.AreEqual(explicitLaunch ? 1 : 0, scope.Books.AddCalls);
                Assert.AreSame(explicitLaunch ? scope.Added : scope.Seed, Get(scope.Fixture, "workbook"));
                Assert.IsFalse(scope.Fixture.PreserveForDiagnosticRecovery);
                string evidence = Path.Combine(scope.Root, "private-unsaved-workbook.json");
                Assert.AreEqual(explicitLaunch, File.Exists(evidence));
                if (explicitLaunch)
                {
                    var row = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(evidence));
                    Assert.AreEqual(desktop, row["Desktop"]); Assert.AreEqual(scope.Descriptor.Id, row["ProcessId"]);
                    Assert.AreEqual("", row["SavedPath"]); Assert.AreEqual(false, row["HelperSaveInvoked"]);
                    Assert.AreEqual(true, row["SeedClosedWithoutSaving"]);
                }
            }
        }

        [DataTestMethod]
        [DataRow(null, null)]
        [DataRow(null, "")]
        [DataRow(null, "relative.jsonl")]
        [DataRow(null, "C:\\trace.jsonl:alternate")]
        [DataRow("Q027Private", null)]
        [DataRow("Q027Private", "")]
        [DataRow("Q027Private", "relative.jsonl")]
        [DataRow("Q027Private", "C:\\trace.jsonl:alternate")]
        public void InvalidExplicitTraceRefusesBothLaunchersBeforeAnyWorkbookMutation(string desktop, string trace)
        {
            int privateGuards = 0, launches = 0;
            Assert.ThrowsException<ArgumentException>(() => ExcelVbeFixture.StartSelectedBootstrap(desktop, "1", trace,
                value => privateGuards++, value => { launches++; return null; }, () => { launches++; return null; }));
            Assert.AreEqual(desktop == null ? 0 : 1, privateGuards); Assert.AreEqual(0, launches);
        }

        [TestMethod]
        public void WrongPrivateDesktopRefusesBeforeTraceValidationOrAnyLauncher()
        {
            var primary = new InvalidOperationException("wrong desktop"); int guards = 0, launches = 0;
            Assert.AreSame(primary, Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.StartSelectedBootstrap("Q027Private", "1", null,
                value => { guards++; throw primary; }, value => { launches++; return null; }, () => { launches++; return null; })));
            Assert.AreEqual(1, guards); Assert.AreEqual(0, launches);
        }

        [DataTestMethod]
        [DataRow(null, "pid")]
        [DataRow(null, "image")]
        [DataRow(null, "birth")]
        [DataRow("Q027Private", "pid")]
        [DataRow("Q027Private", "image")]
        [DataRow("Q027Private", "birth")]
        public void ExistingOwnedIdentityGuardFailureCannotFallbackOrEnterUnsavedTransition(string desktop, string member)
        {
            using (var scope = new Scope())
            {
                int starts = 0, fallback = 0; Exception primary = null;
                var actual = Capture(() => ExcelVbeFixture.StartSelectedBootstrap(desktop, "1", scope.Trace, value => { }, value =>
                {
                    starts++;
                    try
                    {
                        ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(71, @"C:\Qualification\EXCEL.EXE", "exact-birth",
                        member == "pid" ? 72 : 71, member == "image" ? @"C:\Foreign\EXCEL.EXE" : @"C:\Qualification\EXCEL.EXE",
                        member == "birth" ? "foreign-birth" : "exact-birth");
                    }
                    catch (Exception error) { primary = error; throw; }
                    return scope.Fixture;
                }, () => { fallback++; return scope.Fixture; }));
                Assert.IsInstanceOfType(actual, typeof(InvalidOperationException)); Assert.AreSame(primary, actual);
                Assert.AreEqual(1, starts); Assert.AreEqual(0, fallback); Assert.AreEqual(0, scope.Calls.Count);
                Assert.AreSame(scope.Seed, Get(scope.Fixture, "workbook"));
            }
        }

        [DataTestMethod, DataRow(null), DataRow("Q027Private")]
        public void ExplicitLauncherFailureIsNotRetriedOrConvertedToComActivation(string desktop)
        {
            using (var scope = new Scope())
            {
                int starts = 0, fallback = 0;
                Assert.AreSame(scope.Primary, Assert.ThrowsException<IOException>(() => ExcelVbeFixture.StartSelectedBootstrap(desktop, "1", scope.Trace,
                    value => { }, value => { starts++; throw scope.Primary; }, () => { fallback++; return scope.Fixture; })));
                Assert.AreEqual(1, starts); Assert.AreEqual(0, fallback); Assert.AreEqual(0, scope.Calls.Count);
            }
        }

        [DataTestMethod]
        [DataRow(null, "close-before")]
        [DataRow(null, "close-after")]
        [DataRow(null, "add-before")]
        [DataRow(null, "add-after")]
        [DataRow(null, "count-error")]
        [DataRow(null, "extra-workbook")]
        [DataRow(null, "path-error")]
        [DataRow(null, "saved-path")]
        [DataRow(null, "name-error")]
        [DataRow(null, "evidence")]
        [DataRow("Q027Private", "close-before")]
        [DataRow("Q027Private", "close-after")]
        [DataRow("Q027Private", "add-before")]
        [DataRow("Q027Private", "add-after")]
        [DataRow("Q027Private", "count-error")]
        [DataRow("Q027Private", "extra-workbook")]
        [DataRow("Q027Private", "path-error")]
        [DataRow("Q027Private", "saved-path")]
        [DataRow("Q027Private", "name-error")]
        [DataRow("Q027Private", "evidence")]
        public void UncertainSeedAndUnsavedTransitionsRetainTheSameFixtureWithoutCleanupOrReplay(string desktop, string fault)
        {
            using (var scope = new Scope())
            {
                scope.Fault = fault;
                if (fault == "evidence") Directory.CreateDirectory(Path.Combine(scope.Root, "private-unsaved-workbook.json"));
                Exception failure = Capture(() => scope.Start(desktop, "1")); Assert.IsNotNull(failure);
                if (fault != "extra-workbook" && fault != "saved-path" && fault != "evidence") Assert.AreSame(scope.Primary, failure);
                Assert.IsTrue(scope.Fixture.PreserveForDiagnosticRecovery);
                var retained = (IList)typeof(ExcelVbeFixture).GetField("retainedBootstraps", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                lock (retained) Assert.IsTrue(retained.Contains(scope.Fixture));
                Assert.AreEqual(1, scope.Seed.CloseCalls); Assert.AreEqual(fault.StartsWith("close-", StringComparison.Ordinal) ? 0 : 1, scope.Books.AddCalls);
                Assert.AreEqual(fault != "close-before", scope.Seed.Closed);
                Assert.AreEqual(!fault.StartsWith("close-", StringComparison.Ordinal) && fault != "add-before", scope.Books.Added);
                Assert.AreSame(fault.StartsWith("close-", StringComparison.Ordinal) ? scope.Seed : fault.StartsWith("add-", StringComparison.Ordinal) ? null : scope.Added,
                    Get(scope.Fixture, "workbook"));
                int calls = scope.Calls.Count;
                Assert.ThrowsException<InvalidOperationException>(scope.Fixture.Dispose);
                Assert.AreEqual(calls, scope.Calls.Count); Assert.IsFalse(scope.Calls.Contains("Quit"));
                Assert.AreSame(scope.Descriptor, Get(scope.Fixture, "ownedProcess"));
            }
        }
    }
}
