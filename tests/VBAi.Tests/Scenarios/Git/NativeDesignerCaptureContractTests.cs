using System;
using VBAi.Tests.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Scenarios
{
    /// <summary>Pure ownership/identity checks for the native designer screenshot helper; no Windows API is called.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeDesignerCaptureContractTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void VerifiedDesignerOrRootFallbackKeepsExactOwnerAndActiveIdentity(bool zeroDesignerHandle)
        {
            var state = State();
            if (zeroDesignerHandle) state.DesignerHandle = 0;
            var target = ExcelVbeFixture.SelectGitDesignerCaptureTarget(state, 42, _ => 42);
            Assert.AreEqual(new IntPtr(zeroDesignerHandle ? 88 : 44), target);
        }

        [DataTestMethod]
        [DataRow("project")]
        [DataRow("designer")]
        [DataRow("hidden-designer")]
        [DataRow("hidden-root")]
        [DataRow("designer-type")]
        [DataRow("active-type")]
        [DataRow("caption")]
        [DataRow("empty-caption")]
        [DataRow("root-zero")]
        [DataRow("root-owner")]
        [DataRow("zero-pid")]
        public void RootFallbackRefusesUnqualifiedIdentityVisibilityTypeCaptionOrOwner(string scenario)
        {
            var state = State(); state.DesignerHandle = 0;
            uint expectedPid = 42;
            switch (scenario)
            {
                case "project": state.ProjectIdentityMatches = false; break;
                case "designer": state.DesignerIdentityMatches = false; break;
                case "hidden-designer": state.DesignerVisible = false; break;
                case "hidden-root": state.MainVisible = false; break;
                case "designer-type": state.DesignerType = 0; break;
                case "active-type": state.ActiveType = 0; break;
                case "caption": state.ActiveCaption = "Other form"; break;
                case "empty-caption": state.ActiveCaption = state.DesignerCaption = ""; break;
                case "root-zero": state.MainHandle = 0; break;
                case "zero-pid": expectedPid = 0; break;
            }
            Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.SelectGitDesignerCaptureTarget(state,
                expectedPid, _ => scenario == "root-owner" ? 7u : 42u));
        }

        [TestMethod]
        public void ForeignNonzeroDesignerHandleDoesNotFallBackToAnOwnedRoot()
        {
            var state = State();
            Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.SelectGitDesignerCaptureTarget(state,
                42, handle => handle == new IntPtr(88) ? 42u : 7u));
        }

        private static ExcelVbeFixture.GitDesignerCaptureState State()
        {
            return new ExcelVbeFixture.GitDesignerCaptureState {
                DesignerHandle = 44, MainHandle = 88, DesignerCaption = "QualificationForm", ActiveCaption = "QualificationForm",
                DesignerType = 1, ActiveType = 1, DesignerVisible = true, MainVisible = true,
                ProjectIdentityMatches = true, DesignerIdentityMatches = true
            };
        }
    }
}
