namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.IO;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed class LoadLogTests
    {
        [TestMethod]
        public void LoadDiagnosticsIncludeTimestampAndTolerateFilesystemFailures()
        {
            var previous = LoadLog.AppendText;
            try
            {
                string captured = null;
                LoadLog.AppendText = (path, text) => { Assert.AreEqual(LoadLog.PathName, path); captured = text; };
                LoadLog.Write("Disposable load diagnostic");
                StringAssert.Contains(captured, "Disposable load diagnostic" + Environment.NewLine);
                Assert.IsTrue(DateTime.TryParse(captured.Split(' ')[0], out _));
                LoadLog.AppendText = (path, text) => { throw new IOException("Disposable locked file"); }; LoadLog.Write("IO refused");
                LoadLog.AppendText = (path, text) => { throw new UnauthorizedAccessException("Disposable access refusal"); }; LoadLog.Write("Access refused");
            }
            finally { LoadLog.AppendText = previous; }
        }
    }
}
