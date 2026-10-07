using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using System.Threading.Tasks;
namespace VBAi.Tests.Unit
{
    public sealed partial class VbeSessionTests
    {
        [STATestMethod]
        public void NestedPublicationGeneralAuthorizationPermitsOnlyInventoryForOriginalOwner()
        {
            var first = Create(); var other = Create(); var flags = BindingFlags.NonPublic | BindingFlags.Instance;
            var owner = typeof(VbeSession).GetField("macroOwner", BindingFlags.NonPublic | BindingFlags.Static);
            try
            {
                owner.SetValue(null, first.Session);
                foreach (string name in new[] { "macroInFlight", "generalInFlight" }) typeof(VbeSession).GetField(name, flags).SetValue(first.Session, true);
                typeof(VbeSession).GetField("macroAuthorizationDepth", flags).SetValue(first.Session, 1);
                Assert.IsFalse(first.Session.Execute(new Request { Command = "list_projects" }).Ok, "Both nested authorization scopes are required.");
                typeof(VbeSession).GetField("generalAuthorizationDepth", flags).SetValue(first.Session, 1);
                Assert.IsTrue(first.Session.Execute(new Request { Command = "list_projects" }).Ok);
                Assert.IsFalse(first.Session.Execute(new Request { Command = "create_standalone_project" }).Ok);
                Assert.IsFalse(other.Session.Execute(new Request { Command = "list_projects" }).Ok);
                Assert.ThrowsException<InvalidOperationException>(() => other.Session.RequireGeneralSettled());
                typeof(VbeSession).GetField("macroQuarantined", flags).SetValue(first.Session, true);
                Assert.IsFalse(first.Session.Execute(new Request { Command = "list_projects" }).Ok);
                Assert.IsTrue(first.Session.Execute(new Request { Command = "status" }).Ok);
            }
            finally { owner.SetValue(null, null); }
        }
        [STATestMethod]
        public void QuarantinedOwningStaBlocksEverySessionEvenAfterOwnerIsReleased()
        {
            var first = Create(); var other = Create(); var field = typeof(VbeSession).GetField("macroOwnerQuarantined", BindingFlags.NonPublic | BindingFlags.Static);
            bool previous = (bool)field.GetValue(null);
            try
            {
                field.SetValue(null, true);
                foreach (var session in new[] { first.Session, other.Session })
                {
                    Assert.IsFalse(session.Execute(new Request { Command = "list_projects" }).Ok);
                    Assert.IsFalse(session.Execute(new Request { Command = "create_standalone_project" }).Ok);
                    Assert.ThrowsException<InvalidOperationException>(() => session.RequireGeneralSettled());
                    Assert.IsTrue(session.Execute(new Request { Command = "status" }).Ok);
                }
            }
            finally { field.SetValue(null, previous); }
        }
        [STATestMethod]
        public void NestedGeneralReadCannotBeInvokedOutsideOriginalPublication()
        {
            var fixture = Create(); var method = typeof(VbeSession).GetMethod("ReadPublicationGeneralAsync", BindingFlags.NonPublic | BindingFlags.Instance);
            int authorized = 0;
            var pending = (Task<object>)method.Invoke(fixture.Session, new object[] { new Request { Command = "read_project_general", Project = "VBAProject", ExpectedProjectVersion = "supplied", ExpectedMode = 2 }, new Action<bool>(_ => authorized++), new Action(() => { }) });
            Assert.ThrowsException<InvalidOperationException>(() => pending.GetAwaiter().GetResult());
            Assert.AreEqual(0, authorized, "No project selection or native General read may precede admission.");
        }
    }
}
