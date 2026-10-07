using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Proves bounded same-thread read observation and preservation of terminal failures.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeObservationReadTests
    {
        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5)]
        public void GetterReturnsOnFirstAcceptedReadOnTheCallingThread(int rejections)
        {
            int calls = 0, thread = Thread.CurrentThread.ManagedThreadId;
            var waits = new List<int>(); var records = new List<int>();
            string result = OfficeObservationRead.Getter(() =>
            {
                Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId);
                if (calls++ < rejections) throw new COMException("rejected", unchecked((int)0x80010001));
                return "native read";
            }, (attempt, delay) => { records.Add(attempt); Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId); }, waits.Add);
            Assert.AreEqual("native read", result); Assert.AreEqual(rejections + 1, calls);
            CollectionAssert.AreEqual(new List<int>(new[] { 50, 100, 200, 400, 800 }).GetRange(0, rejections), waits);
            Assert.AreEqual(rejections, records.Count);
        }

        [TestMethod]
        public void ExhaustionThrowsTheOriginalLastGetterErrorAndNeverWaitsAgain()
        {
            int calls = 0, waits = 0;
            var error = new COMException("terminal rejection", unchecked((int)0x80010001));
            Assert.AreSame(error, Assert.ThrowsException<COMException>(() => OfficeObservationRead.Getter<int>(
                () => { calls++; throw error; }, pause: delay => waits++)));
            Assert.AreEqual(6, calls); Assert.AreEqual(5, waits);
        }

        [DataTestMethod]
        [DataRow(unchecked((int)0x80010108)), DataRow(unchecked((int)0x80004005))]
        public void OtherComErrorsAreNeverRetried(int hresult)
        {
            int calls = 0;
            var error = new COMException("other", hresult);
            Assert.AreSame(error, Assert.ThrowsException<COMException>(() => OfficeObservationRead.Getter<int>(
                () => { calls++; throw error; }, pause: delay => Assert.Fail("Unexpected wait."))));
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public void ManagedAndEvidenceFailuresAreNotSwallowed()
        {
            var error = new InvalidOperationException("original");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => OfficeObservationRead.Getter<int>(() => { throw error; })));
            int calls = 0;
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => OfficeObservationRead.Getter<int>(
                () => { calls++; throw new COMException("busy", unchecked((int)0x80010001)); },
                (attempt, delay) => { throw error; }, delay => Assert.Fail("No wait after lost evidence."))));
            Assert.AreEqual(1, calls);
            Assert.ThrowsException<ArgumentNullException>(() => OfficeObservationRead.Getter<int>(null));
        }
    }
}
