using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks historical reference isolation without any provider or native project access.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptReferencesTests
    {
        private static VbeChatReference Reference(string revision = "old") => new VbeChatReference
        { Project = "Project1", Module = "Module1", Name = "Calculate", Kind = "Sub", Sha256 = revision, StartLine = 3, EndLine = 8 };

        /// <summary>Later turns, missing context and legacy boundaries must not retarget an older message.</summary>
        [DataTestMethod]
        [DataRow("explicit")]
        [DataRow("turn")]
        [DataRow("interleaved")]
        [DataRow("legacy")]
        [DataRow("explicit-empty")]
        [DataRow("legacy-empty")]
        [DataRow("missing-turn")]
        [DataRow("foreign-entry")]
        [DataRow("foreign-explicit")]
        [DataRow("user-empty")]
        [DataRow("null-entry")]
        [DataRow("null-transcript")]
        [DataRow("structured-legacy-boundary")]
        public void CapturedContextDoesNotBorrowFutureOrForeignTurns(string scenario)
        {
            var original = Reference(); var later = Reference("new");
            var prompt = new ChatEntry { Speaker = "Vous", TurnId = "first", References = new[] { original } };
            var answer = new ChatEntry { Speaker = "Assistant", TurnId = "first" };
            var future = new ChatEntry { Speaker = "Vous", TurnId = "second", References = new[] { later } };
            var entries = new List<ChatEntry> { prompt, answer, future };
            bool expected = true;
            if (scenario == "explicit") answer.References = new[] { original };
            if (scenario == "interleaved") entries.Insert(1, new ChatEntry { Speaker = "Vous", TurnId = "other", References = new[] { later } });
            if (scenario == "legacy") { prompt.TurnId = null; answer.TurnId = null; future.TurnId = null; }
            if (scenario == "explicit-empty") { answer.References = new VbeChatReference[0]; expected = false; }
            if (scenario == "legacy-empty") { prompt.TurnId = null; answer.TurnId = null; prompt.References = new VbeChatReference[0]; expected = false; }
            if (scenario == "missing-turn") { answer.TurnId = "unknown"; expected = false; }
            if (scenario == "foreign-entry") { entries.Remove(answer); expected = false; }
            if (scenario == "foreign-explicit") { entries.Remove(answer); answer.References = new[] { original }; }
            if (scenario == "user-empty") { answer.Speaker = "Vous"; expected = false; }
            if (scenario == "null-entry") { answer = null; expected = false; }
            if (scenario == "null-transcript") { entries = null; expected = false; }
            if (scenario == "structured-legacy-boundary") { answer.TurnId = null; expected = false; }
            var links = ChatTranscriptReferences.Resolve(entries, answer);
            Assert.AreEqual(expected ? 1 : 0, links.Count, scenario);
            if (expected) Assert.AreSame(original, links[original.Token], scenario);
        }

        /// <summary>Conflicting revisions or source ranges cannot select an arbitrary destination.</summary>
        [DataTestMethod]
        [DataRow("revision")]
        [DataRow("start")]
        [DataRow("end")]
        [DataRow("procedure-kind")]
        [DataRow("kind")]
        public void ConflictingSnapshotsWithTheSameTokenAreNotClickable(string field)
        {
            var first = Reference(); var conflict = Reference();
            if (field == "revision") conflict.Sha256 = "new";
            if (field == "start") conflict.StartLine++;
            if (field == "end") conflict.EndLine++;
            if (field == "procedure-kind") conflict.ProcKind++;
            if (field == "kind") conflict.Kind = "Function";
            var unrelated = new VbeChatReference { Project = "Other", Kind = "Projet" };
            var entry = new ChatEntry { References = new[] { first, conflict, first, unrelated } };
            var links = ChatTranscriptReferences.Resolve(null, entry);
            Assert.AreEqual(1, links.Count);
            Assert.AreSame(unrelated, links[unrelated.Token]);
            Assert.IsFalse(links.ContainsKey(first.Token));
        }

        /// <summary>Equivalent duplicates coalesce; corrupt entries do not hide unrelated valid attachments.</summary>
        [TestMethod]
        public void EquivalentDuplicatesAndNullOrBlankProjectsAreHandledWithoutMutatingTheSnapshot()
        {
            var first = Reference();
            var entry = new ChatEntry { References = new[] { null, new VbeChatReference(), new VbeChatReference { Project = " " }, first, Reference() } };
            var links = ChatTranscriptReferences.Resolve(null, entry);
            Assert.AreEqual(1, links.Count); Assert.AreSame(first, links[first.Token]);
            links.Clear();
            Assert.AreEqual(5, entry.References.Length);
            Assert.AreEqual(1, ChatTranscriptReferences.Resolve(null, entry).Count);
        }

        /// <summary>Loading or forking a serialized conversation preserves earlier targets independently of later attachments.</summary>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ReloadedAndForkedConversationKeepsItsOriginalReferenceRevision(bool legacy)
        {
            var old = Reference();
            var entries = new[]
            {
                new ChatEntry { Speaker = "Vous", TurnId = legacy ? null : "first", References = new[] { old } },
                new ChatEntry { Speaker = "Assistant", TurnId = legacy ? null : "first", Text = old.Token },
                new ChatEntry { Speaker = "Vous", TurnId = legacy ? null : "second", References = new[] { Reference("new") } }
            };
            var serializer = new JavaScriptSerializer();
            var loaded = serializer.Deserialize<ChatEntry[]>(serializer.Serialize(entries));
            Assert.AreEqual("old", ChatTranscriptReferences.Resolve(loaded, loaded[1])[old.Token].Sha256);
            var fork = new[] { loaded[0], loaded[1] };
            Assert.AreEqual("old", ChatTranscriptReferences.Resolve(fork, fork[1])[old.Token].Sha256);
            Assert.AreEqual("new", ChatTranscriptReferences.Resolve(loaded, loaded[2])[old.Token].Sha256);
        }

        /// <summary>Current token matching remains ordinal; unrelated case variants do not replace one another.</summary>
        [TestMethod]
        public void CaseVariantsKeepSeparateCapturedTokens()
        {
            var first = Reference(); var second = Reference("new"); second.Project = "project1";
            var entry = new ChatEntry { References = new[] { first, second } };
            var links = ChatTranscriptReferences.Resolve(null, entry);
            Assert.AreEqual(2, links.Count); Assert.AreSame(first, links[first.Token]); Assert.AreSame(second, links[second.Token]);
        }
    }
}
