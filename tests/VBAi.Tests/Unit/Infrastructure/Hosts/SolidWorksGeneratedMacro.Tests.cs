using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Security.Cryptography;
using System.Text;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Fixture admission scenarios for the observed 2019 template; never launches SOLIDWORKS.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SolidWorksGeneratedMacroTests
    {
        private const string Source = "Dim swApp As Object\r\nSub main()\r\n\r\nSet swApp = Application.SldWorks\r\nEnd Sub";
        private static string Hash(string text)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }

        /// <summary>Known generated code admits only a hash-guarded full-module replacement, never execution.</summary>
        [DataTestMethod, DataRow("CRLF"), DataRow("LF"), DataRow("case"), DataRow("indent")]
        public void KnownGeneratedDefaultProducesOneGuardedBlankingRequest(string variation)
        {
            string code = variation == "LF" ? Source.Replace("\r\n", "\n") : variation == "case" ? Source.ToUpperInvariant() :
                variation == "indent" ? "  " + Source.Replace("\r\n", "\r\n  ") : Source;
            var request = SolidWorksGeneratedMacro.PlanBlankModule(@"C:\Disposable\fresh.swp", "Module1", 1, 2, code, Hash(code), 5);
            Assert.AreEqual("replace_lines", request.Command);
            Assert.AreEqual(@"C:\Disposable\fresh.swp", request.Project);
            Assert.AreEqual("Module1", request.Module);
            Assert.AreEqual(2, request.ExpectedMode);
            Assert.AreEqual(Hash(code), request.ExpectedSha256);
            Assert.AreEqual(1, request.StartLine);
            Assert.AreEqual(5, request.Count);
            Assert.AreEqual("Option Explicit\r\n", request.Text);
        }

        /// <summary>An already blank standard module needs no mutation.</summary>
        [DataTestMethod, DataRow(""), DataRow("\r\n"), DataRow("Option Explicit"), DataRow("  option explicit\r\n")]
        public void BlankSourceProducesNoEdit(string code)
        {
            Assert.IsNull(SolidWorksGeneratedMacro.PlanBlankModule(@"C:\Disposable\fresh.swp", "Module1", 1, 2, code, Hash(code), 2));
        }

        /// <summary>Every change outside the exact generated template is preserved by refusal.</summary>
        [DataTestMethod]
        [DataRow("comment"), DataRow("extra procedure"), DataRow("application"), DataRow("declaration"), DataRow("code")]
        public void ArbitrarySourceIsNeverBlanked(string variation)
        {
            string code = variation == "comment" ? Source + "\r\n' personal note" :
                variation == "extra procedure" ? Source + "\r\nSub UserCode()\r\nEnd Sub" :
                variation == "application" ? Source.Replace("Application.SldWorks", "Application.Other") :
                variation == "declaration" ? Source.Replace("Dim swApp As Object", "Dim swApp As String") : "MsgBox 1";
            Assert.ThrowsException<InvalidOperationException>(() => SolidWorksGeneratedMacro.PlanBlankModule(
                @"C:\Disposable\fresh.swp", "Module1", 1, 2, code, Hash(code), 10));
        }

        /// <summary>Invalid target, mode, component, checksum and line metadata cannot produce a mutation.</summary>
        [DataTestMethod]
        [DataRow("path"), DataRow("extension"), DataRow("module"), DataRow("class"), DataRow("form")]
        [DataRow("native"), DataRow("run"), DataRow("break"), DataRow("checksum"), DataRow("short count")]
        [DataRow("negative count"), DataRow("large count"), DataRow("null source")]
        public void InvalidAdmissionMetadataIsRefused(string fault)
        {
            string path = fault == "path" ? "fresh.swp" : fault == "extension" ? @"C:\Disposable\fresh.bas" : @"C:\Disposable\fresh.swp";
            int type = fault == "class" ? 2 : fault == "form" ? 3 : fault == "native" ? 100 : 1;
            int mode = fault == "run" ? 0 : fault == "break" ? 1 : 2;
            int count = fault == "short count" ? 4 : fault == "negative count" ? -1 : fault == "large count" ? 129 : 5;
            Assert.ThrowsException<InvalidOperationException>(() => SolidWorksGeneratedMacro.PlanBlankModule(path,
                fault == "module" ? "" : "Module1", type, mode, fault == "null source" ? null : Source,
                fault == "checksum" ? new string('0', 64) : Hash(Source), count));
        }
    }
}
