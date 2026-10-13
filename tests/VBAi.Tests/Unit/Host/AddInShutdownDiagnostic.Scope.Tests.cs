using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    /// <summary>Verifies attribution to the exact innermost active invocation during reentrant cleanup.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class AddInShutdownDiagnosticScopeTests
    {
        [DataTestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void AncestorCannotBorrowAnActiveChildsExceptionAndResumesItsOwnObservationAfterChildCompletion(bool grandchild, bool com)
        {
            var observation = new ShutdownDiagnosticTestData.Observation();
            var root = observation.Begin(); root.Enter("ExistingParentOperation");
            var child = root.Child("Dispose"); child.Enter("ExistingChildOperation");
            var active = child;
            if (grandchild) { active = child.Child("Dispose"); active.Enter("ExistingGrandchildOperation"); }
            Exception error = com ? (Exception)new COMException("caught exact child operation") : new IOException("caught exact child operation");
            int before = observation.Events.Count;
            root.ExceptionObserved(error);
            if (grandchild) child.ExceptionObserved(error);
            Assert.AreEqual(before, observation.Events.Count, "An ancestor must not attribute its active child's exception to its own operation.");
            Assert.IsFalse(root.DiagnosticFailed); Assert.IsFalse(active.DiagnosticFailed);
            active.ExceptionObserved(error);
            Assert.AreEqual(before + 1, observation.Events.Count);
            var row = (Dictionary<string, object>)observation.Events.Last();
            Assert.AreEqual(active.InvocationId, row["InvocationId"]);
            Assert.AreEqual(grandchild ? "ExistingGrandchildOperation" : "ExistingChildOperation", row["Stage"]);
            Assert.AreEqual("ExceptionObserved", row["Phase"]); Assert.IsNull(row["FaultCaught"]);
            if (grandchild) { active.Returned("ExistingGrandchildOperation"); active.Complete(); }
            child.Returned("ExistingChildOperation"); child.Complete();
            before = observation.Events.Count; root.ExceptionObserved(error);
            Assert.AreEqual(before + 1, observation.Events.Count);
            row = (Dictionary<string, object>)observation.Events.Last();
            Assert.AreEqual(root.InvocationId, row["InvocationId"]); Assert.AreEqual("ExistingParentOperation", row["Stage"]);
            root.Returned("ExistingParentOperation"); root.Complete();
            Assert.IsFalse(root.DiagnosticFailed);
        }
    }
}
