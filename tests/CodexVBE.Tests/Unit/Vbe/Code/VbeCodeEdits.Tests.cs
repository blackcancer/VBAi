using System;
using System.Collections;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie le budget, les versions et les échecs de l’historique de code géré.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeCodeEditsTests
    {
        /// <summary>Contenu public lisible par la frontière dynamique de production.</summary>
        public sealed class Snapshot
        {
            /// <summary>Source simulée courante.</summary>
            public string Code { get; set; }
        }
        /// <summary>Calcule l’empreinte du contenu de la fixture.</summary>
        private static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
        /// <summary>Accède à la pile privée pour contrôler les limites de conservation.</summary>
        private static IList Stack(VbeCodeEdits edits, string name)
        { return (IList)typeof(VbeCodeEdits).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(edits); }

        /// <summary>Élimine les entrées anciennes selon le nombre et le volume de texte.</summary>
        [TestMethod]
        public void HistoryBudgetsAndNoChangeRecordingKeepBoundedSessionState()
        {
            var edits = new VbeCodeEdits(_ => Response.Failure("unused"));
            edits.Record("p", "m", "same", "same"); Assert.AreEqual(0, Stack(edits, "undo").Count);
            for (int index = 0; index < 51; index++) edits.Record("p", "m", "before", "after" + index);
            Assert.AreEqual(50, Stack(edits, "undo").Count);
            edits.Record("p", "m", new string('x', 2097152), new string('y', 2097153));
            Assert.AreEqual(0, Stack(edits, "undo").Count);
            edits.Record("p", "m", "one", "two"); Assert.AreEqual(1, Stack(edits, "undo").Count);
        }

        /// <summary>Refuse lectures, écritures et relectures invalides sans consommer une entrée de replay.</summary>
        [TestMethod]
        public void ReadWriteReadbackAndRevisionFailuresPreserveHistory()
        {
            string code = "one"; bool failRead = false, failWrite = false, wrongReadback = false;
            var edits = new VbeCodeEdits(command => {
                if (command.Command == "read_module") return failRead ? Response.Failure("read failed") : Response.Success(new Snapshot { Code = code });
                if (failWrite) return Response.Failure("write failed");
                code = wrongReadback ? "different" : command.Text;
                return Response.Success("written");
            });
            var request = new Request { Project = "P", Module = "M", StartLine = 1, Count = 1, Action = "replace", Query = "one", Text = "two", ExpectedSha256 = Hash(code) };
            failRead = true; StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => edits.Edit(request, false)).Message, "read failed"); failRead = false;
            foreach (string sha in new[] { null, " ", "stale" })
            { request.ExpectedSha256 = sha; Assert.ThrowsException<InvalidOperationException>(() => edits.Edit(request, true)); }
            request.ExpectedSha256 = Hash(code); failWrite = true;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => edits.Edit(request, false)).Message, "write failed"); failWrite = false;
            wrongReadback = true; Assert.ThrowsException<InvalidOperationException>(() => edits.Edit(request, false)); wrongReadback = false;
            code = "two"; edits.Record("other", "M", "one", "two"); edits.Record("P", "other", "one", "two"); edits.Record("P", "M", "one", "two");
            request.ExpectedSha256 = Hash(code); failWrite = true;
            Assert.ThrowsException<InvalidOperationException>(() => edits.Replay(request, false)); Assert.AreEqual(3, Stack(edits, "undo").Count);
            failWrite = false; edits.Replay(request, false); Assert.AreEqual("one", code);
            request.ExpectedSha256 = Hash(code); edits.Replay(request, true); Assert.AreEqual("two", code);
            Assert.AreEqual(0, Stack(edits, "redo").Count);
            request.Module = "missing"; Assert.ThrowsException<InvalidOperationException>(() => edits.Replay(request, false));
        }
    }
}
