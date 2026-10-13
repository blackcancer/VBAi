using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks dispatcher selection and the actual MSTest invocation across the managed tool fixture boundary.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ToolBoundaryTestClassAttributeTests
    {
        /// <summary>Retains ordinary method display names while selecting the isolated managed dispatcher.</summary>
        /// <param name="kind">Ordinary, data-row, or unspecified method metadata.</param>
        [DataTestMethod, DataRow("ordinary"), DataRow("data"), DataRow("null")]
        public void OrdinaryToolFixturesSelectTheManagedDispatcher(string kind)
        {
            TestMethodAttribute original = kind == "ordinary" ? new TestMethodAttribute("Named fixture") :
                kind == "data" ? new DataTestMethodAttribute() : null;
            var selected = new ToolBoundaryTestClassAttribute().GetTestMethodAttribute(original);
            Assert.IsNotNull(selected);
            Assert.AreNotSame(original, selected);
            if (original != null) Assert.AreEqual(original.DisplayName, selected.DisplayName);
        }

        /// <summary>Retains the original dispatcher for explicit native STA and WinForms fixtures.</summary>
        /// <param name="winForms">Chooses the pumped WinForms dispatcher instead of MSTest's STA dispatcher.</param>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ExplicitUiDispatchersArePreserved(bool winForms)
        {
            TestMethodAttribute original = winForms ? (TestMethodAttribute)new WinFormsTestMethodAttribute() : new STATestMethodAttribute();
            Assert.AreSame(original, new ToolBoundaryTestClassAttribute().GetTestMethodAttribute(original));
        }

        /// <summary>Verifies the executing thread, invocation arguments, cultures and result identity.</summary>
        /// <param name="ambientContext">Installs a caller context to verify that it is not borrowed by the managed fixture.</param>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ManagedInvocationUsesMtaAndRetainsArgumentsCultureAndResult(bool ambientContext)
        {
            var previous = SynchronizationContext.Current;
            int owner = Thread.CurrentThread.ManagedThreadId;
            var context = ambientContext ? new SynchronizationContext() : null;
            var culture = CultureInfo.CurrentCulture;
            var uiCulture = CultureInfo.CurrentUICulture;
            object[] args = { new object() };
            var result = new TestResult { Outcome = UnitTestOutcome.Passed, LogOutput = "Retained diagnostic" };
            var invocation = new Invocation(false, args, values =>
            {
                Assert.AreNotEqual(owner, Thread.CurrentThread.ManagedThreadId);
                Assert.AreEqual(ApartmentState.MTA, Thread.CurrentThread.GetApartmentState());
                Assert.IsNull(SynchronizationContext.Current);
                Assert.AreEqual(culture, CultureInfo.CurrentCulture);
                Assert.AreEqual(uiCulture, CultureInfo.CurrentUICulture);
                Assert.IsNull(values, "The original MSTest dispatcher invokes the retained row with no argument override.");
                return result;
            });
            try
            {
                SynchronizationContext.SetSynchronizationContext(context);
                var rows = Selected().Execute(invocation);
                Assert.AreEqual(1, invocation.Calls);
                Assert.AreSame(args, invocation.Arguments);
                Assert.AreSame(result, rows[0]);
                Assert.AreSame(context, SynchronizationContext.Current);
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Prevents a fixture's unpumped UI context from contaminating the next asynchronous invocation.</summary>
        /// <param name="winForms">Uses a real WinForms context instead of a generic caller context.</param>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ManagedInvocationsDoNotShareAnUnpumpedFixtureContext(bool winForms)
        {
            var callerContext = SynchronizationContext.Current;
            int firstThread = 0;
            var first = new Invocation(false, new object[0], values =>
            {
                firstThread = Thread.CurrentThread.ManagedThreadId;
                using (var ownedContext = winForms ? new WindowsFormsSynchronizationContext() : null)
                {
                    SynchronizationContext.SetSynchronizationContext(ownedContext ?? new SynchronizationContext());
                    return new TestResult { Outcome = UnitTestOutcome.Passed };
                }
            });
            var second = new Invocation(false, new object[0], values =>
            {
                Assert.AreNotEqual(firstThread, Thread.CurrentThread.ManagedThreadId);
                Assert.IsNull(SynchronizationContext.Current);
                RequireOrdinaryContinuation().GetAwaiter().GetResult();
                return new TestResult { Outcome = UnitTestOutcome.Passed };
            });
            Assert.AreEqual(UnitTestOutcome.Passed, Selected().Execute(first)[0].Outcome);
            Assert.AreSame(callerContext, SynchronizationContext.Current);
            Assert.AreEqual(UnitTestOutcome.Passed, Selected().Execute(second)[0].Outcome);
            Assert.AreSame(callerContext, SynchronizationContext.Current);
            Assert.AreEqual(1, first.Calls);
            Assert.AreEqual(1, second.Calls);
        }

        /// <summary>Checks that a genuine asynchronous suspension resumes without borrowing an abandoned UI context.</summary>
        private static async Task RequireOrdinaryContinuation()
        {
            await Task.Delay(1);
            Assert.IsNull(SynchronizationContext.Current);
        }

        /// <summary>Honors legacy methods that declare STA as well as ordinary method metadata.</summary>
        /// <param name="data">Selects the ordinary data dispatcher instead of the ordinary method dispatcher.</param>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DeclaredStaMethodIsNeverMovedToTheManagedWorker(bool data)
        {
            int owner = Thread.CurrentThread.ManagedThreadId;
            bool callerIsSta = Thread.CurrentThread.GetApartmentState() == ApartmentState.STA;
            var invocation = new Invocation(true, new object[0], values =>
            {
                Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                if (callerIsSta) Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                else Assert.AreNotEqual(owner, Thread.CurrentThread.ManagedThreadId);
                return new TestResult { Outcome = UnitTestOutcome.Passed };
            });
            new ToolBoundaryTestClassAttribute().GetTestMethodAttribute(data ?
                (TestMethodAttribute)new DataTestMethodAttribute() : new TestMethodAttribute()).Execute(invocation);
            Assert.AreEqual(1, invocation.Calls);
        }

        /// <summary>Retains original failures instead of translating them into passing result rows.</summary>
        /// <param name="thrown">Throws the original exception rather than returning a failing MSTest result.</param>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ManagedFailureAndThrownExceptionArePreserved(bool thrown)
        {
            var error = new InvalidOperationException("Original fixture failure");
            var result = new TestResult { Outcome = UnitTestOutcome.Failed, TestFailureException = error };
            var invocation = new Invocation(false, new object[0], values =>
            {
                if (thrown) throw error;
                return result;
            });
            if (thrown) Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => Selected().Execute(invocation)));
            else Assert.AreSame(result, Selected().Execute(invocation)[0]);
            Assert.AreEqual(1, invocation.Calls);
        }

        /// <summary>Refuses a missing invocation before creating a fixture thread.</summary>
        [TestMethod]
        public void MissingInvocationIsRefused()
            => Assert.ThrowsException<ArgumentNullException>(() => Selected().Execute(null));

        /// <summary>Creates the ordinary managed dispatcher exercised by the tests.</summary>
        /// <returns>An actual selected method attribute.</returns>
        private static TestMethodAttribute Selected()
            => new ToolBoundaryTestClassAttribute().GetTestMethodAttribute(new TestMethodAttribute());

        /// <summary>Supplies a genuine ITestMethod boundary without replacing the dispatcher under test.</summary>
        private sealed class Invocation : ITestMethod
        {
            /// <summary>Callback invoked by the original MSTest method attribute.</summary>
            private readonly Func<object[], TestResult> invoke;
            /// <summary>Counts original MSTest invocations.</summary>
            internal int Calls;
            /// <summary>Creates an invocation with ordinary or explicit STA method metadata.</summary>
            /// <param name="sta">Uses the explicitly attributed STA method.</param>
            /// <param name="arguments">Original row arguments.</param>
            /// <param name="invoke">Callback invoked through ITestMethod.</param>
            internal Invocation(bool sta, object[] arguments, Func<object[], TestResult> invoke)
            {
                MethodInfo = typeof(Metadata).GetMethod(sta ? "StaMetadata" : "OrdinaryMetadata", BindingFlags.Static | BindingFlags.NonPublic);
                Arguments = arguments;
                this.invoke = invoke;
            }
            /// <inheritdoc />
            public string TestMethodName => MethodInfo.Name;
            /// <inheritdoc />
            public string TestClassName => MethodInfo.DeclaringType.FullName;
            /// <inheritdoc />
            public Type ReturnType => MethodInfo.ReturnType;
            /// <inheritdoc />
            public object[] Arguments { get; }
            /// <inheritdoc />
            public ParameterInfo[] ParameterTypes => MethodInfo.GetParameters();
            /// <inheritdoc />
            public MethodInfo MethodInfo { get; }
            /// <inheritdoc />
            public TestResult Invoke(object[] arguments) { Calls++; return invoke(arguments); }
            /// <inheritdoc />
            public Attribute[] GetAllAttributes(bool inherit) => Attribute.GetCustomAttributes(MethodInfo, inherit);
            /// <inheritdoc />
            public T[] GetAttributes<T>(bool inherit) where T : Attribute
                => Array.ConvertAll(MethodInfo.GetCustomAttributes(typeof(T), inherit), value => (T)value);
        }

        /// <summary>Supplies reflected method attributes outside the discovered test class.</summary>
        private static class Metadata
        {
            /// <summary>Ordinary reflected metadata used by the invocation fixture.</summary>
            private static void OrdinaryMetadata() { }
            /// <summary>Explicit STA metadata used to verify dispatcher preservation.</summary>
            [STATestMethod]
            private static void StaMetadata() { }
        }
    }
}
