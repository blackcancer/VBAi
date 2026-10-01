using System;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class PathVisibilityDiagnosticTests
    {
        private const string Local = @"C:\SyntheticLocal", Temp = @"C:\SyntheticTemp";
        private const string Child = "0123456789abcdef0123456789abcdef";
        private static string Manifest(string local, string temp, object version = null) =>
            new JavaScriptSerializer().Serialize(new { Version = version ?? 1, LocalAppData = local, Temp = temp });

        [TestMethod]
        public void ExactGuidChildrenProduceOnlyTheTwoFixedSyntheticNames()
        {
            var paths = PathVisibilityDiagnostic.ValidateManifest(Manifest(Local + "\\" + Child, Temp + "\\" + Child), Local, Temp);
            CollectionAssert.AreEqual(new[] { Local + "\\" + Child, Local + "\\" + Child + "\\" + PathVisibilityDiagnostic.SyntheticName,
                Temp + "\\" + Child, Temp + "\\" + Child + "\\" + PathVisibilityDiagnostic.SyntheticName }, paths);
        }

        [DataTestMethod]
        [DataRow(@"C:\Unrelated\0123456789abcdef0123456789abcdef")]
        [DataRow(@"C:\SyntheticLocal\nested\0123456789abcdef0123456789abcdef")]
        [DataRow(@"C:\SyntheticLocal\not-a-guid")]
        [DataRow(@"C:\SyntheticLocal\0123456789abcdef0123456789abcdef\..\0123456789abcdef0123456789abcdef")]
        [DataRow(@"C:\SyntheticLocal\0123456789abcdef0123456789abcdef:stream")]
        [DataRow(@"\\server\share\0123456789abcdef0123456789abcdef")]
        [DataRow(@"\\?\C:\SyntheticLocal\0123456789abcdef0123456789abcdef")]
        public void NoncanonicalOrOutOfAllowlistPathsAreRefused(string path)
        {
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.ValidateManifest(Manifest(path, Temp + "\\" + Child), Local, Temp));
        }

        [TestMethod]
        public void BadVersionExtraFieldsOversizeAndDuplicatedDirectoriesAreRefused()
        {
            string valid = Manifest(Local + "\\" + Child, Temp + "\\" + Child);
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.ValidateManifest(valid.TrimEnd('}') + ",\"Path\":\"private\"}", Local, Temp));
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.ValidateManifest(Manifest(Local + "\\" + Child, Temp + "\\" + Child, "1"), Local, Temp));
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.ValidateManifest(new string(' ', 4097), Local, Temp));
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.ValidateManifest(Manifest(Local + "\\" + Child, Local + "\\" + Child), Local, Local));
        }

        [TestMethod]
        public void ParameterFreeRequestRefusesAnyAdditionalPathOrProjectField()
        {
            PathVisibilityDiagnostic.RequireParameterFree("{\"Command\":\"diagnostic_path_visibility\"}");
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.RequireParameterFree("{\"Command\":\"diagnostic_path_visibility\",\"Path\":\"private\"}"));
            Assert.ThrowsException<ArgumentException>(() => PathVisibilityDiagnostic.RequireParameterFree("{\"Command\":\"diagnostic_path_visibility\",\"Project\":null}"));
        }

        [TestMethod]
        public void OwnerRequiresExactProcessNativeThreadAndStaWithoutSubstitutingManagedThread()
        {
            PathVisibilityDiagnostic.RequireOwner(42, 7, 42, 7, ApartmentState.STA);
            Assert.ThrowsException<InvalidOperationException>(() => PathVisibilityDiagnostic.RequireOwner(42, 7, 43, 7, ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => PathVisibilityDiagnostic.RequireOwner(42, 7, 42, 8, ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => PathVisibilityDiagnostic.RequireOwner(42, 7, 42, 7, ApartmentState.MTA));
            Assert.ThrowsException<InvalidOperationException>(() => PathVisibilityDiagnostic.RequireOwner(0, 0, 0, 0, ApartmentState.STA));
        }
    }
}
