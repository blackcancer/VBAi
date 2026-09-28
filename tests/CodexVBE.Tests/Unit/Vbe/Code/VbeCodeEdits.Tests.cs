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
    public sealed partial class VbeCodeEditsTests
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

        [TestMethod]
        public void LocalRenameResolvesExactProcedureRevisionAndRequiresLiveDesignModeBeforeWriting()
        {
            const string source = "Sub Run()\nDim value As Long\nDebug.Print value\nEnd Sub";
            string current = source; bool catalogFailure = false, stateFailure = false; int writes = 0, mode = 2;
            var match = new ProcedureRow { Name = "Run", Kind = 0, BodyLine = 1, EndLine = 4 };
            var catalog = new ProcedureCatalog { Sha256 = Hash(source), Procedures = new[] { match } };
            var edits = new VbeCodeEdits(command => {
                switch (command.Command)
                {
                    case "read_module": return Response.Success(new Snapshot { Code = current });
                    case "list_procedures": return catalogFailure ? Response.Failure("catalog unavailable") : Response.Success(catalog);
                    case "debug_state": return stateFailure ? Response.Failure("state unavailable") : Response.Success(new DesignState { Mode = mode });
                    case "replace_lines": current = command.Text; writes++; return Response.Success("written");
                    default: throw new InvalidOperationException(command.Command);
                }
            });
            var request = new Request { Project = "P", Module = "M", Procedure = "Run", ProcKind = 0,
                Query = "value", NewName = "amount", StartLine = 2, StartColumn = 5, ExpectedSha256 = Hash(source), ExpectedMode = 2 };
            catalogFailure = true; StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, true)).Message, "catalog unavailable"); catalogFailure = false;
            catalog.Sha256 = "stale"; Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, true)); catalog.Sha256 = Hash(source).ToUpperInvariant();
            catalog.Procedures = new[] { new ProcedureRow { Name = "Run", Kind = 1 }, new ProcedureRow { Name = "Other", Kind = 0 } };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, true)).Message, "absent");
            catalog.Procedures = new[] { match, match };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, true)).Message, "ambiguous");
            catalog.Procedures = new[] { new ProcedureRow { Name = "Other", Kind = 0 }, match, new ProcedureRow { Name = "Run", Kind = 1 } };
            dynamic preview = edits.RenameLocal(request, true);
            Assert.AreEqual(source, (string)preview.Before); StringAssert.Contains((string)preview.After, "Dim amount"); Assert.IsTrue((bool)preview.Changed);
            Assert.AreEqual(0, writes); Assert.AreEqual(source, current);
            request.ExpectedMode = 1; Assert.ThrowsException<ArgumentException>(() => edits.RenameLocal(request, false)); request.ExpectedMode = 2;
            stateFailure = true; Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, false)); stateFailure = false;
            mode = 1; Assert.ThrowsException<InvalidOperationException>(() => edits.RenameLocal(request, false)); mode = 2;
            Assert.AreEqual("written", edits.RenameLocal(request, false)); Assert.AreEqual(1, writes); StringAssert.Contains(current, "Debug.Print amount");
        }
    }
}
