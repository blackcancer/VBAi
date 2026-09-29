using System;
using System.IO;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;
using VBAi.Tests.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Matrice du transport NamedPipe : connexion retry, émission unique, perte de réponse et délai.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeBridgeClientTests
    {
        /// <summary>Une commande de mutation traverse une seule fois le pipe unique et la réponse est relue.</summary>
        [TestMethod]
        public void SuccessfulOwnedPipeReceivesExactlyOneSerializedRequest()
        {
            using (var pipe = new BridgeClientPipeFixture("{\"Ok\":true,\"Data\":42}"))
            {
                Assert.IsTrue(pipe.Ready.Wait(3000));
                var response = VbeBridgeClient.Read(pipe.Name, new { Command = "mutate", Text = "line\r\nnext" }, 2000);
                Assert.AreEqual(true, response["Ok"]); Assert.AreEqual(42, response["Data"]); Assert.AreEqual(1, pipe.Count);
                var request = VbeBridgeClient.Object(new JavaScriptSerializer().DeserializeObject(pipe.Payload));
                Assert.AreEqual("mutate", request["Command"]); Assert.AreEqual("line\r\nnext", request["Text"]); Assert.IsNull(pipe.Failure);
            }
        }

        /// <summary>Le serveur démarre après un véritable timeout de connexion, avant toute émission.</summary>
        [TestMethod]
        public void TransientConnectTimeoutRetriesConnectionAndEmitsOnce()
        {
            string name = "VBAi.Tests." + Guid.NewGuid().ToString("N"); BridgeClientPipeFixture pipe = null; int retries = 0;
            try
            {
                var response = VbeBridgeClient.Read(name, new { Command = "mutate" }, 2000, 30, 3, 0, attempt => {
                    retries++; Assert.AreEqual(1, attempt);
                    pipe = new BridgeClientPipeFixture("{\"Ok\":true}", pipeName: name); Assert.IsTrue(pipe.Ready.Wait(3000));
                });
                Assert.AreEqual(true, response["Ok"]); Assert.AreEqual(1, retries); Assert.IsNotNull(pipe); Assert.AreEqual(1, pipe.Count);
            }
            finally { pipe?.Dispose(); }
        }

        /// <summary>Une perte de réponse après réception ne réémet jamais la mutation.</summary>
        [TestMethod]
        public void DisconnectAfterSendReportsUncertaintyWithoutReconnect()
        {
            using (var pipe = new BridgeClientPipeFixture(null, disconnect: true))
            {
                Assert.IsTrue(pipe.Ready.Wait(3000)); int retries = 0;
                var failure = Assert.ThrowsException<IOException>(() => VbeBridgeClient.Read(pipe.Name, new { Command = "mutate" }, 2000,
                    onConnectionRetry: attempt => retries++));
                StringAssert.Contains(failure.Message, "not retried"); Assert.IsTrue(pipe.Received.Wait(3000)); Assert.AreEqual(1, pipe.Count); Assert.AreEqual(0, retries);
            }
        }

        /// <summary>Le délai de réponse configuré ferme le client sans livrer une seconde requête.</summary>
        [TestMethod]
        public void ResponseDeadlineAfterSendEmitsOnceAndNeverRetries()
        {
            using (var pipe = new BridgeClientPipeFixture(null, hold: true))
            {
                Assert.IsTrue(pipe.Ready.Wait(3000)); int retries = 0;
                var failure = Assert.ThrowsException<TimeoutException>(() => VbeBridgeClient.Read(pipe.Name, new { Command = "mutate" }, 100,
                    onConnectionRetry: attempt => retries++));
                StringAssert.Contains(failure.Message, "delivery is uncertain"); Assert.IsTrue(pipe.Received.Wait(3000)); Assert.AreEqual(1, pipe.Count); Assert.AreEqual(0, retries);
            }
        }

        /// <summary>Un JSON malformé ou scalaire est une erreur de réponse après une seule émission.</summary>
        [TestMethod]
        public void MalformedAndNonObjectResponsesAreNeverRetried()
        {
            foreach (string reply in new[] { "{broken", "42" })
                using (var pipe = new BridgeClientPipeFixture(reply))
                {
                    Assert.IsTrue(pipe.Ready.Wait(3000)); int retries = 0;
                    if (reply == "42") Assert.ThrowsException<InvalidDataException>(() => VbeBridgeClient.Read(pipe.Name, new { Command = "mutate" }, 2000,
                        onConnectionRetry: attempt => retries++));
                    else Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read(pipe.Name, new { Command = "mutate" }, 2000,
                        onConnectionRetry: attempt => retries++));
                    Assert.AreEqual(1, pipe.Count); Assert.AreEqual(0, retries);
                }
        }

        /// <summary>Sans serveur les seules tentatives de connexion s'épuisent sans requête.</summary>
        [TestMethod]
        public void MissingServerReturnsNullAfterBoundedConnectionAttempts()
        {
            int retries = 0;
            Assert.IsNull(VbeBridgeClient.Read("VBAi.Tests." + Guid.NewGuid().ToString("N"), new { Command = "mutate" }, 100, 10, 2, 0,
                attempt => retries++)); Assert.AreEqual(1, retries);
        }

        /// <summary>Les overloads historiques ciblent le nom de pipe du seul identifiant fourni.</summary>
        [TestMethod]
        public void ExistingProcessAndCommandOverloadsPreservePipeAndPayload()
        {
            foreach (bool textual in new[] { false, true })
            {
                int id = int.MaxValue - (Guid.NewGuid().GetHashCode() & 0xfffff);
                using (var pipe = new BridgeClientPipeFixture("{\"Ok\":true}", pipeName: "VBAi." + id))
                {
                    Assert.IsTrue(pipe.Ready.Wait(3000));
                    var response = textual ? VbeBridgeClient.Read(id, "status") : VbeBridgeClient.Read(id, new { Command = "status" });
                    Assert.AreEqual(true, response["Ok"]); Assert.AreEqual(1, pipe.Count); StringAssert.Contains(pipe.Payload, "status");
                }
            }
        }

        /// <summary>Les bornes invalides sont refusées avant de connecter ou émettre.</summary>
        [TestMethod]
        public void InvalidTimeoutAndRetryBoundsAreRejected()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read((string)null, new { Command = "mutate" }));
            Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read(" ", new { Command = "mutate" }));
            foreach (int value in new[] { 0, 600001 }) Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read("owned", null, value));
            foreach (int value in new[] { 0, 10001 }) Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read("owned", null, 100, value));
            foreach (int value in new[] { 0, 101 }) Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read("owned", null, 100, 10, value));
            foreach (int value in new[] { -1, 10001 }) Assert.ThrowsException<ArgumentException>(() => VbeBridgeClient.Read("owned", null, 100, 10, 1, value));
        }
    }
}