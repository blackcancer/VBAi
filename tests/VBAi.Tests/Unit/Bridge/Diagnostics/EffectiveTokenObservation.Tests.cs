using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EffectiveTokenObservationTests
    {
        private sealed class Reader : EffectiveTokenObservation.ITokenReader
        {
            internal bool ThreadToken, PrimaryToken = true, MetadataFails, CloseFails;
            internal int ThreadError = EffectiveTokenObservation.ErrorNoToken, PrimaryError = 5, PrimaryCalls, MetadataCalls, Closes;
            internal IDictionary<string, object> Values = new Dictionary<string, object> { ["TokenType"] = 1 };
            public bool OpenThread(out IntPtr token, out int error) { token = ThreadToken ? new IntPtr(7) : IntPtr.Zero; error = ThreadError; return ThreadToken; }
            public bool OpenPrimary(out IntPtr token, out int error) { PrimaryCalls++; token = PrimaryToken ? new IntPtr(8) : IntPtr.Zero; error = PrimaryError; return PrimaryToken; }
            public IDictionary<string, object> Metadata(IntPtr token) { MetadataCalls++; if (MetadataFails) throw new InvalidOperationException("metadata"); return Values; }
            public void Close(IntPtr token) { Closes++; if (CloseFails) throw new InvalidOperationException("close"); }
        }

        [TestMethod]
        public void EffectiveThreadTokenNeverFallsBackToPrimaryAndClosesOnce()
        {
            var reader = new Reader { ThreadToken = true };
            var result = EffectiveTokenObservation.Read(reader);
            Assert.AreEqual("Thread", result["Source"]); Assert.AreEqual("READ", result["State"]);
            Assert.AreEqual(0, reader.PrimaryCalls); Assert.AreEqual(1, reader.MetadataCalls); Assert.AreEqual(1, reader.Closes);
        }
        [TestMethod]
        public void ErrorNoTokenIsTheOnlyAllowedFallbackAndItsEvidenceIsRetained()
        {
            var reader = new Reader(); var result = EffectiveTokenObservation.Read(reader);
            Assert.AreEqual("PrimaryAfterErrorNoToken", result["Source"]); Assert.AreEqual(1008, result["OpenThreadTokenError"]);
            Assert.AreEqual(1, reader.PrimaryCalls); Assert.AreEqual(1, reader.Closes);
        }
        [DataTestMethod, DataRow(5), DataRow(87), DataRow(0)]
        public void AnyOtherThreadOpenErrorRefusesPrimaryFallback(int error)
        {
            var reader = new Reader { ThreadError = error }; var result = EffectiveTokenObservation.Read(reader);
            Assert.AreEqual("UNVERIFIED", result["State"]); Assert.AreEqual(error, result["OpenThreadTokenError"]);
            Assert.AreEqual(0, reader.PrimaryCalls); Assert.AreEqual(0, reader.MetadataCalls); Assert.AreEqual(0, reader.Closes);
        }
        [TestMethod]
        public void PrimaryOpenFailureDoesNotQueryOrCloseAnUnopenedToken()
        {
            var reader = new Reader { PrimaryToken = false }; var result = EffectiveTokenObservation.Read(reader);
            Assert.AreEqual(5, result["OpenProcessTokenError"]); Assert.AreEqual(0, reader.MetadataCalls); Assert.AreEqual(0, reader.Closes);
        }
        [TestMethod]
        public void MetadataAndCloseFailuresArePreservedIndependentlyWithoutAnotherOpen()
        {
            var reader = new Reader { ThreadToken = true, MetadataFails = true, CloseFails = true };
            var result = EffectiveTokenObservation.Read(reader);
            Assert.AreEqual(typeof(InvalidOperationException).FullName, result["ErrorType"]);
            Assert.AreEqual(typeof(InvalidOperationException).FullName, result["CloseErrorType"]);
            Assert.AreEqual("PARTIAL", result["State"]); Assert.AreEqual(1, reader.Closes); Assert.AreEqual(0, reader.PrimaryCalls);
        }
        [TestMethod]
        public void PartialNativeMetadataCannotBecomeReadAcceptance()
        {
            var reader = new Reader { Values = new Dictionary<string, object> { ["MetadataErrors"] = new Dictionary<string, int> { ["IntegritySid"] = 5 } } };
            Assert.AreEqual("PARTIAL", EffectiveTokenObservation.Read(reader)["State"]);
            Assert.AreEqual(1, reader.Closes);
        }
    }
}
