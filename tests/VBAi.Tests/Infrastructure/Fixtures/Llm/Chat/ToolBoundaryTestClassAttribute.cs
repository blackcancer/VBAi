using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace VBAi.Tests.Unit
{
    /// <summary>Runs managed tool fixtures on MTA while retaining explicitly selected UI test dispatchers.</summary>
    [AttributeUsage(AttributeTargets.Class)]
    internal sealed class ToolBoundaryTestClassAttribute : TestClassAttribute
    {
        /// <summary>Separates managed boundary tests from the native STA fixtures in the same class.</summary>
        /// <param name="testMethodAttribute">Method dispatcher selected by MSTest, including custom UI dispatchers.</param>
        /// <returns>The original explicit dispatcher, or an isolated MTA dispatcher for an ordinary managed test.</returns>
        public override TestMethodAttribute GetTestMethodAttribute(TestMethodAttribute testMethodAttribute)
        {
            var original = testMethodAttribute ?? new TestMethodAttribute();
            var type = original.GetType();
            if (type != typeof(TestMethodAttribute) && type != typeof(DataTestMethodAttribute)) return original;
            return new ManagedMethod(original);
        }

        /// <summary>Runs only ordinary managed fixtures on a separate MTA thread with their original MSTest invocation.</summary>
        private sealed class ManagedMethod : TestMethodAttribute
        {
            /// <summary>Original dispatcher retaining MSTest's test invocation and result handling.</summary>
            private readonly TestMethodAttribute original;

            /// <summary>Retains the selected method dispatcher and its display name.</summary>
            /// <param name="original">Ordinary method dispatcher selected by MSTest.</param>
            internal ManagedMethod(TestMethodAttribute original) : base(original.DisplayName) { this.original = original; }

            /// <summary>Invokes the managed fixture on MTA without moving an explicitly attributed STA test.</summary>
            /// <param name="testMethod">MSTest invocation including its original arguments and initialization pipeline.</param>
            /// <returns>The original result objects, including their failures and diagnostic output.</returns>
            public override TestResult[] Execute(ITestMethod testMethod)
            {
                if (testMethod == null) throw new ArgumentNullException(nameof(testMethod));
                // Some legacy methods declare both TestMethod and STATestMethod.
                // Preserve that explicit STA metadata whichever attribute MSTest selected first.
                var sta = testMethod.MethodInfo?.GetCustomAttributes(typeof(STATestMethodAttribute), true)
                    .Cast<STATestMethodAttribute>().SingleOrDefault();
                if (sta != null && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                    return sta.Execute(testMethod);
                // MSTest handles this marker in its adapter. Direct invocation of Execute
                // inherits TestMethodAttribute and cannot select the apartment for us.
                return ExecuteIsolated(sta ?? original, testMethod,
                    sta == null ? ApartmentState.MTA : ApartmentState.STA);
            }

            /// <summary>Invokes the retained dispatcher on the required apartment without borrowing a caller UI context.</summary>
            /// <param name="dispatcher">Original method dispatcher retaining the MSTest pipeline and results.</param>
            /// <param name="testMethod">Invocation with its original row and initialization metadata.</param>
            /// <param name="apartment">Required apartment for this invocation.</param>
            /// <returns>Original results after the original fixture thread has exited.</returns>
            private static TestResult[] ExecuteIsolated(TestMethodAttribute dispatcher, ITestMethod testMethod, ApartmentState apartment)
            {
                TestResult[] results = null;
                Exception failure = null;
                var culture = CultureInfo.CurrentCulture;
                var uiCulture = CultureInfo.CurrentUICulture;
                var thread = new Thread(() =>
                {
                    CultureInfo.CurrentCulture = culture;
                    CultureInfo.CurrentUICulture = uiCulture;
                    SynchronizationContext.SetSynchronizationContext(null);
                    try { results = dispatcher.Execute(testMethod); }
                    catch (Exception error) { failure = error; }
                }) { IsBackground = true, Name = "VBAi managed tool fixture" };
                thread.SetApartmentState(apartment);
                thread.Start();
                if (!thread.Join(TimeSpan.FromMinutes(2)))
                {
                    // Preserve the original deadline failure without advancing into another
                    // fixture while this exact original thread still has admitted work.
                    thread.Join();
                    throw new TimeoutException("The managed tool fixture did not finish on its isolated " + apartment + " thread.");
                }
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
                return results;
            }
        }
    }
}
