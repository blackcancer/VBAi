namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Exécute le scénario avec une boucle de messages WinForms sur son propre thread STA.</summary>
    internal sealed class WinFormsTestMethodAttribute : TestMethodAttribute
    {
        /// <summary>Démarre la boucle native, exécute le test puis termine le thread et relaie son résultat.</summary>
        /// <param name="testMethod">Scénario découvert par MSTest.</param>
        /// <returns>Résultats du scénario exécuté sur le thread UI.</returns>
        public override TestResult[] Execute(ITestMethod testMethod)
        {
            TestResult[] results = null;
            Exception failure = null;
            var thread = new Thread(() => {
                using (var dispatcher = new Control())
                {
                    dispatcher.CreateControl();
                    dispatcher.BeginInvoke(new Action(() => {
                        try { results = base.Execute(testMethod); }
                        catch (Exception ex) { failure = ex; }
                        finally { Application.ExitThread(); }
                    }));
                    Application.Run();
                }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(TimeSpan.FromMinutes(2))) throw new TimeoutException("Disposable WinForms test thread timed out.");
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            return results;
        }
    }
}
