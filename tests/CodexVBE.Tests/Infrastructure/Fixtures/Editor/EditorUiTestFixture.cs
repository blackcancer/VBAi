using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Remonte les erreurs de boucle WinForms aux tests au lieu de bloquer sur le dialogue JIT.</summary>
    public abstract class EditorUiTestFixture
    {
        [ThreadStatic] private static Exception uiFailure;

        [TestInitialize]
        public void CaptureEditorUiFailures()
        {
            uiFailure = null;
            Application.ThreadException += OnUiFailure;
        }

        private static void OnUiFailure(object sender, ThreadExceptionEventArgs args)
        {
            if (uiFailure == null) uiFailure = args.Exception;
        }

        internal static void ThrowIfUiFailed()
        {
            if (uiFailure != null)
                throw new AssertFailedException("Unhandled editor UI exception: " + uiFailure);
        }


        [TestCleanup]
        public void ValidateEditorUiFailures()
        {
            Application.ThreadException -= OnUiFailure;
            var failure = uiFailure;
            uiFailure = null;
            if (failure != null)
                throw new AssertFailedException("Unhandled editor UI exception: " + failure);
        }
    }
}
