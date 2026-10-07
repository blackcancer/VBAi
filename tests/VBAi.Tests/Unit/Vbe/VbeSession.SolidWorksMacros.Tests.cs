using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeSessionTests
    {
        [STATestMethod]
        public void NativeMacroOwnerExcludesOtherSessionsOnTheSameOwningSta()
        {
            var first = Create(); var second = Create();
            var owner = typeof(VbeSession).GetField("macroOwner", BindingFlags.NonPublic | BindingFlags.Static);
            try
            {
                owner.SetValue(null, first.Session);
                Assert.IsFalse(second.Session.Execute(new Request { Command = "list_projects" }).Ok);
                Assert.ThrowsException<InvalidOperationException>(() => second.Session.RequireGeneralSettled());
                Assert.IsTrue(second.Session.Execute(new Request { Command = "status" }).Ok);
                typeof(VbeSession).GetField("macroAuthorizationDepth", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(first.Session, 1);
                Assert.IsTrue(first.Session.Execute(new Request { Command = "list_projects" }).Ok);
                Assert.IsFalse(second.Session.Execute(new Request { Command = "create_standalone_project" }).Ok);
            }
            finally { owner.SetValue(null, null); }
            Assert.IsTrue(second.Session.Execute(new Request { Command = "list_projects" }).Ok);
        }

        [DataTestMethod, DataRow("macroInFlight"), DataRow("macroQuarantined")]
        public void PendingOrUncertainNativeMacroBlocksSessionMutationAndAdmission(string field)
        {
            var fixture = Create();
            typeof(VbeSession).GetField(field, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(fixture.Session, true);
            var response = fixture.Session.Execute(new Request { Command = "create_standalone_project", ExpectedMode = 2 });
            Assert.IsFalse(response.Ok);
            Assert.IsTrue(fixture.Session.Execute(new Request { Command = "status" }).Ok);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.RequireGeneralSettled());
        }

        [TestMethod]
        public void NativeMacroAuthorizationAllowsOnlyProjectInventoryInsideItsScope()
        {
            var fixture = Create();
            var flags = BindingFlags.NonPublic | BindingFlags.Instance;
            typeof(VbeSession).GetField("macroInFlight", flags).SetValue(fixture.Session, true);
            Assert.IsFalse(fixture.Session.Execute(new Request { Command = "list_projects" }).Ok);
            typeof(VbeSession).GetField("macroAuthorizationDepth", flags).SetValue(fixture.Session, 1);
            Assert.IsTrue(fixture.Session.Execute(new Request { Command = "list_projects" }).Ok);
            Assert.IsFalse(fixture.Session.Execute(new Request { Command = "create_standalone_project" }).Ok);
            typeof(VbeSession).GetField("macroQuarantined", flags).SetValue(fixture.Session, true);
            Assert.IsFalse(fixture.Session.Execute(new Request { Command = "list_projects" }).Ok);
        }

        [STATestMethod]
        public void NativeMacroRouteRequiresOriginalAuthorizationBeforeAnyBackendEntry()
        {
            var fixture = Create();
            var pending = fixture.Session.SolidWorksMacroAsync(new Request { Command = "create_solidworks_macro", ExpectedMode = 2 });
            Assert.ThrowsException<InvalidOperationException>(() => pending.GetAwaiter().GetResult());
            Assert.IsTrue(fixture.Session.Execute(new Request { Command = "list_projects" }).Ok);
        }
    }
}
