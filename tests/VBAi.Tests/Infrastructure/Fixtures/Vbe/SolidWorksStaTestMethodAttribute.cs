using System;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Reuses the synthetic persistence STA adapter; never launches or attaches SOLIDWORKS.</summary>
    internal sealed class SolidWorksStaTestMethodAttribute : TestMethodAttribute
    {
        public override TestResult[] Execute(ITestMethod testMethod) => ExecuteScenario(() => base.Execute(testMethod));

        /// <summary>Keeps the original MSTest rows/failure while rejecting success with an unfinished task.</summary>
        internal static TestResult[] ExecuteScenario(Func<TestResult[]> original)
        {
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(original);
            if (run.Error != null) ExceptionDispatchInfo.Capture(run.Error).Throw();
            return AccessStaTestMethodAttribute.PreserveResults(run.Value, run.Retained, run.Diagnostic);
        }

        internal static Task<object> StartSave(Func<Task<object>> start) => AccessStaTestMethodAttribute.StartSave(start);
        internal static string Describe(Task task) => AccessStaTestMethodAttribute.Describe(task);
    }
}
