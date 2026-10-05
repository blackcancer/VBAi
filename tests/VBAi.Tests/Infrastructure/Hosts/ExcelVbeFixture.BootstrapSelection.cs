using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Selects the existing owned launch for a private desktop or the exact explicit-bootstrap opt-in.</summary>
        internal static ExcelVbeFixture StartSelectedBootstrap(string desktop, string explicitBootstrap, string tracePath,
            Action<string> requireCurrentDesktop, Func<string, ExcelVbeFixture> startOwned, Func<ExcelVbeFixture> startCom)
        {
            bool privateDesktop = !string.IsNullOrWhiteSpace(desktop);
            if (!privateDesktop && explicitBootstrap != "1") return startCom();
            if (privateDesktop) requireCurrentDesktop(desktop);
            var fixture = startOwned(ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(tracePath));
            // Both desktop paths preserve the generic scenario's original unsaved-workbook
            // precondition. The existing owned bootstrap has already checked native identity
            // and the loaded candidate, and armed any requested shutdown observation.
            fixture.PreserveForDiagnosticRecovery = true;
            try
            {
                ((dynamic)fixture.workbook).Close(false);
                Release(fixture.workbook); fixture.workbook = null;
                fixture.workbook = ((dynamic)fixture.workbooks).Add();
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)fixture.workbooks).Count));
                Assert.IsTrue(string.IsNullOrEmpty(Convert.ToString(((dynamic)fixture.workbook).Path)));
                // Keep the existing evidence filename and members for the private path.
                fixture.WriteEvidence("private-unsaved-workbook.json", new {
                    Desktop = desktop, fixture.ProcessId, Workbook = Convert.ToString(((dynamic)fixture.workbook).Name),
                    SavedPath = Convert.ToString(((dynamic)fixture.workbook).Path), HelperSaveInvoked = false,
                    SeedClosedWithoutSaving = true, Utc = DateTime.UtcNow.ToString("o") });
                fixture.PreserveForDiagnosticRecovery = false;
                return fixture;
            }
            catch
            {
                lock (retainedBootstraps) retainedBootstraps.Add(fixture);
                throw; // Unknown Close/Add outcomes never authorize replay or cleanup.
            }
        }
    }
}
