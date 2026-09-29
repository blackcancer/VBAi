namespace VBAi.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Reflection;
    using System.Text;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ProcessInputTests
    {
        [TestMethod]
        public void NativeChildReceivesUtf8WithoutBomAndConsoleEncodingIsRestoredAfterSuccessAndFailure()
        {
            using(var scope=new LlmBoundaryScope())
            {
                string executable=Compile(scope);
                var encoding=Console.InputEncoding;var field=ProcessInput.InputEncodingField();object previous=field.GetValue(null);
                using(var process=Child(executable))
                {
                    Assert.IsTrue(ProcessInput.StartWithoutPreamble(process));
                    Assert.AreSame(previous,field.GetValue(null));Assert.AreSame(encoding,Console.InputEncoding);
                    process.StandardInput.Write("é\n");process.StandardInput.Close();
                    Assert.AreEqual("C3-A9-0A",process.StandardOutput.ReadToEnd());Assert.IsTrue(process.WaitForExit(5000));Assert.AreEqual(0,process.ExitCode);
                }
                using(var missing=Child(executable+".missing"))
                    Assert.ThrowsException<Win32Exception>(()=>ProcessInput.StartWithoutPreamble(missing));
                Assert.AreSame(previous,field.GetValue(null));
            }
        }

        [TestMethod]
        public void MissingRuntimeEncodingFieldFailsBeforeStartingAnyProcess()
        {
            var original=ProcessInput.InputEncodingField;
            try
            {
                ProcessInput.InputEncodingField=()=>null;
                using(var process=new Process())
                {
                    var error=Assert.ThrowsException<InvalidOperationException>(()=>ProcessInput.StartWithoutPreamble(process));
                    StringAssert.Contains(error.Message,".NET Framework");Assert.ThrowsException<InvalidOperationException>(()=> {var id=process.Id;});
                }
            }
            finally {ProcessInput.InputEncodingField=original;}
        }
    }
}
