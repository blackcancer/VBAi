namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using System.Security.Cryptography;
    using System.Text;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeProcedureMutationTests
    {
        [TestMethod]
        public void MutationRejectsEveryInvalidVbideRangeWithoutEditing()
        {
            foreach (int operation in new[] { 0, 1, 2 })
                foreach (int fault in new[] { 0, 1, 2, 3 })
                {
                    var f = new Fixture("Private Sub Run()\r\nEnd Sub");
                    if (fault == 0) { f.Module.BodyOverride = 1; f.Module.StartOverride = 2; }
                    if (fault == 1) { f.Module.BodyOverride = 0; f.Module.StartOverride = 0; }
                    if (fault == 2) f.Module.CountOverride = 0;
                    if (fault == 3) f.Module.CountOverride = 99;
                    var request = f.Request("Private Sub Run()\r\nDebug.Print 1\r\nEnd Sub");
                    Assert.ThrowsException<InvalidOperationException>(() =>
                    {
                        if (operation == 0) f.Navigation.ReplaceProcedure(request);
                        else if (operation == 1) f.Navigation.RemoveProcedure(request);
                        else f.Navigation.Procedures("Projet", "Module1");
                    });
                    Assert.AreEqual("Private Sub Run()\r\nEnd Sub", f.Module.Code);
                }
        }

        [TestMethod]
        public void ProcedureCreationValidatesReadbackAndRollsBackItsInsertion()
        {
            foreach (int fault in new[] { 0, 1, 2, 3, 4 })
            {
                var f = new Fixture("");
                if (fault == 0) f.Module.BodyOverride = 0;
                if (fault == 1) f.Module.KindOverride = 3;
                if (fault == 2) f.Module.IdentityOverride = "Wrong";
                if (fault == 3) f.Module.IgnoreInsert = true;
                if (fault == 4) f.Module.FailNextInsert = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.CreateProcedure(f.Request("Sub Run()\r\nEnd Sub")));
                Assert.AreEqual("", f.Module.Code, "Creation rollback fault " + fault);
            }
        }

        [TestMethod]
        public void ReplacementAndRemovalKeepOriginalCodeAfterReadbackOrWriteFailures()
        {
            foreach (bool remove in new[] { false, true })
                foreach (int fault in new[] { 0, 1, 2, 3 })
                {
                    var f = new Fixture("Sub Run()\r\nEnd Sub");
                    if (fault == 0) f.Module.KindOverride = 3;
                    if (fault == 1) f.Module.IdentityOverride = "Wrong";
                    if (fault == 2) f.Module.FailDelete = true;
                    if (fault == 3)
                    {
                        if (remove) f.Module.IgnoreDelete = true;
                        else f.Module.FailNextInsert = true;
                    }
                    var request = f.Request("Sub Run()\r\nDebug.Print 1\r\nEnd Sub");
                    Assert.ThrowsException<InvalidOperationException>(() =>
                    {
                        if (remove) f.Navigation.RemoveProcedure(request);
                        else f.Navigation.ReplaceProcedure(request);
                    });
                    if (!(remove && fault == 3)) Assert.AreEqual("Sub Run()\r\nEnd Sub", f.Module.Code);
                }
        }

        [TestMethod]
        public void RollbackFailuresPreserveTheOriginalExceptionAndRequireInspection()
        {
            foreach (bool remove in new[] { false, true })
            {
                var f = new Fixture("Sub Run()\r\nEnd Sub");
                if (remove) f.Module.PretendStillPresentAfterDelete = true;
                f.Module.FailEveryInsert = true;
                var request = f.Request("Sub Run()\r\nDebug.Print 1\r\nEnd Sub");
                var error = Assert.ThrowsException<InvalidOperationException>(() =>
                {
                    if (remove) f.Navigation.RemoveProcedure(request);
                    else f.Navigation.ReplaceProcedure(request);
                });
                StringAssert.Contains(error.Message, "rollback needs inspection");
                Assert.IsNotNull(error.InnerException);
            }
            var corrupt = new Fixture("Sub Run()\r\nEnd Sub");
            corrupt.Module.FailNextInsert = true;
            corrupt.Module.CorruptRollback = true;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() =>
                corrupt.Navigation.ReplaceProcedure(corrupt.Request("Sub Run()\r\nDebug.Print 1\r\nEnd Sub"))).Message, "rollback needs inspection");
        }

        [TestMethod]
        public void MissingEndStatementsAndUnavailableModulesFailBeforeMutation()
        {
            foreach (bool remove in new[] { false, true })
            {
                var f = new Fixture("Sub Run()\r\nDebug.Print 1");
                var r = f.Request("Sub Run()\r\nEnd Sub");
                Assert.ThrowsException<InvalidOperationException>(() => { if (remove) f.Navigation.RemoveProcedure(r); else f.Navigation.ReplaceProcedure(r); });
                Assert.AreEqual("Sub Run()\r\nDebug.Print 1", f.Module.Code);
                r.Module = null;
                Assert.ThrowsException<ArgumentException>(() => f.Navigation.CreateProcedure(r));
                r.Module = "Absent";
                Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.CreateProcedure(r));
            }
        }

        [TestMethod]
        public void RemovingProceduresRequiresValidIdentifiersKindsAndHashes()
        {
            var f = new Fixture("Sub Run()\r\nEnd Sub");
            Assert.ThrowsException<ArgumentException>(() => f.Navigation.RemoveProcedure(null));
            foreach (int fault in new[] { 0, 1, 2, 3, 4 })
            {
                var r = f.Request(null);
                if (fault == 0) r.Procedure = null;
                if (fault == 1) r.Procedure = "0invalid";
                if (fault == 2) r.ProcKind = -1;
                if (fault == 3) r.ProcKind = 4;
                if (fault == 4) r.ExpectedSha256 = null;
                Assert.ThrowsException<ArgumentException>(() => f.Navigation.RemoveProcedure(r));
            }
            foreach (bool remove in new[] { false, true })
            {
                var r = f.Request("Sub Run()\r\nEnd Sub");
                f.Project.Mode = 1;
                Assert.ThrowsException<InvalidOperationException>(() => { if (remove) f.Navigation.RemoveProcedure(r); else f.Navigation.ReplaceProcedure(r); });
                f.Project.Mode = 2;
                r.ExpectedSha256 = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => { if (remove) f.Navigation.RemoveProcedure(r); else f.Navigation.ReplaceProcedure(r); });
            }
        }

        [TestMethod]
        public void ClassPropertiesRoundTripForEveryAccessorKind()
        {
            foreach (int kind in new[] { 1, 2, 3 })
            {
                var f = new Fixture("", 2);
                string accessor = new[] { "", "Let", "Set", "Get" }[kind];
                var r = f.Request("Public Property " + accessor + " Run()\r\nEnd Property");
                r.ProcKind = kind;
                dynamic made = f.Navigation.CreateProcedure(r);
                Assert.AreEqual(kind, (int)made.ProcKind);
                dynamic listed = f.Navigation.Procedures("Projet", "Module1");
                Assert.AreEqual(kind, (int)listed.Procedures[0].Kind);
                r.ExpectedSha256 = made.Sha256;
                r.Text = "Public Property " + accessor + " Run()\r\nDebug.Print 1\r\nEnd Property";
                dynamic replaced = f.Navigation.ReplaceProcedure(r);
                r.ExpectedSha256 = replaced.Sha256;
                f.Navigation.RemoveProcedure(r);
                Assert.AreEqual("", f.Module.Code);
            }
        }

        [TestMethod]
        public void EventCreationValidatesNamesModesTargetsVersionsAndReadback()
        {
            for (int fault = 0; fault < 14; fault++)
            {
                var f = new Fixture("", 3);
                dynamic tree = f.Forms.Tree("Projet", "Module1");
                var r = f.Request(null);
                r.Form = "Module1"; r.ObjectName = "UserForm"; r.EventName = "Click"; r.ExpectedTreeVersion = tree.TreeVersion;
                if (fault == 0) r.EventName = null;
                if (fault == 1) r.EventName = "0Click";
                if (fault == 2) r.ObjectName = null;
                if (fault == 3) r.ObjectName = "0Object";
                if (fault == 4) r.ExpectedSha256 = null;
                if (fault == 5) r.ExpectedTreeVersion = null;
                if (fault == 6) f.Project.Mode = 1;
                if (fault == 7) r.Form = "Absent";
                if (fault == 8) f.Component.Type = 1;
                if (fault == 9) r.ExpectedSha256 = "stale";
                if (fault == 10) f.Module.EventBodyOverride = 0;
                if (fault == 11) f.Module.KindOverride = 3;
                if (fault == 12) f.Module.IdentityOverride = "Wrong";
                if (fault == 13) f.Module.IgnoreInsert = true;
                if (fault < 6) Assert.ThrowsException<ArgumentException>(() => f.Navigation.CreateEventProcedure(r));
                else Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.CreateEventProcedure(r));
            }
            var duplicate = new Fixture("Private Sub UserForm_Click()\r\nEnd Sub", 3);
            dynamic duplicateTree = duplicate.Forms.Tree("Projet", "Module1");
            var request = duplicate.Request(null); request.Form = "Module1"; request.ObjectName = "UserForm"; request.EventName = "Click"; request.ExpectedTreeVersion = duplicateTree.TreeVersion;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => duplicate.Navigation.CreateEventProcedure(request)).Message, "already exists");
        }

        [TestMethod]
        public void SourceInsertionValidatesAllInputsAndFileRaceOutcomes()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-source-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "source.bas");
                File.WriteAllText(path, "'source", new UTF8Encoding(false));
                var f = new Fixture("");
                Assert.ThrowsException<ArgumentException>(() => f.Navigation.InsertCodeFile(null));
                for (int fault = 0; fault < 5; fault++)
                {
                    var r = f.Request(null); r.StartLine = 1; r.Path = path;
                    if (fault == 0) r.ExpectedSha256 = null;
                    if (fault == 1) r.StartLine = 0;
                    if (fault == 2) r.Path = null;
                    if (fault == 3) r.Path = "relative.bas";
                    if (fault == 4) r.Path = "C:relative.bas";
                    Assert.ThrowsException<ArgumentException>(() => f.Navigation.InsertCodeFile(r));
                }
                var request = f.Request(null); request.Path = path; request.StartLine = 1;
                File.Delete(path);
                Assert.ThrowsException<FileNotFoundException>(() => f.Navigation.InsertCodeFile(request));
                foreach (int size in new[] { 0, 262145 })
                {
                    File.WriteAllBytes(path, new byte[size]);
                    Assert.ThrowsException<ArgumentException>(() => f.Navigation.InsertCodeFile(request));
                }
                File.WriteAllText(path, "'source", new UTF8Encoding(false));
                foreach (int size in new[] { 0, 262145 })
                {
                    f.Navigation.ReadSourceBytes = _ => new byte[size];
                    Assert.ThrowsException<ArgumentException>(() => f.Navigation.InsertCodeFile(request));
                    Assert.ThrowsException<ArgumentException>(() => f.Navigation.InspectCodeFile(path));
                }
                f.Navigation.ReadSourceBytes = File.ReadAllBytes;
                foreach (string source in new[] { "\0", "   " })
                {
                    File.WriteAllText(path, source, new UTF8Encoding(false));
                    Assert.ThrowsException<ArgumentException>(() => f.Navigation.InsertCodeFile(request));
                }
                File.WriteAllText(path, "'source", new UTF8Encoding(false));
                f.Project.Mode = 1;
                Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.InsertCodeFile(request));
                f.Project.Mode = 2; request.Module = "absent";
                Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.InsertCodeFile(request));
                request.Module = "Module1"; request.StartLine = 2;
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Navigation.InsertCodeFile(request));
            }
            finally { Directory.Delete(directory, true); }
        }

        [TestMethod]
        public void SourceInsertionRestoresExactTextOrReportsFailedRollback()
        {
            string path = Path.Combine(Path.GetTempPath(), "VBAi-source-" + Guid.NewGuid().ToString("N") + ".bas");
            File.WriteAllText(path, "'été", new UTF8Encoding(true));
            try
            {
                for (int fault = 0; fault < 5; fault++)
                {
                    var f = new Fixture("'original");
                    var r = f.Request(null); r.Path = path; r.StartLine = 2;
                    if (fault == 0) f.Module.IgnoreInsert = true;
                    if (fault == 1) f.Module.FailNextInsert = true;
                    if (fault >= 2) f.Module.ReadLinesOverride = (start, count, actual) => f.Module.InsertCalls > 0 && start == 2 ? "'ete" : actual;
                    if (fault == 3) f.Module.FailDelete = true;
                    if (fault == 4) f.Module.ReadLinesOverride = (start, count, actual) => f.Module.DeleteCalls > 0 ? "corrupted rollback" : f.Module.InsertCalls > 0 && start == 2 ? "'ete" : actual;
                    var error = Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.InsertCodeFile(r));
                    if (fault >= 3) { StringAssert.Contains(error.Message, "rollback needs inspection"); Assert.IsNotNull(error.InnerException); }
                    else Assert.AreEqual("'original", f.Module.Code);
                }
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void ReadbackRejectsUnchangedTextEvenWhenVbideReportsNewLines()
        {
            var f = new Fixture("'original");
            f.Module.ReadLinesOverride = (start, count, actual) => f.Module.InsertCalls > 0 && start == 1 ? "'original" : actual;
            Assert.ThrowsException<InvalidOperationException>(() => f.Navigation.CreateProcedure(f.Request("Sub Run()\r\nEnd Sub")));
            Assert.AreEqual("'original", f.Module.Code);
            var removal = new Fixture("Sub Run()\r\nEnd Sub\r\n'last");
            string before = removal.Module.Code;
            removal.Module.ReadLinesOverride = (start, count, actual) => removal.Module.DeleteCalls > 0 && start == 1 ? before : actual;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => removal.Navigation.RemoveProcedure(removal.Request(null))).Message, "did not change");
            Assert.AreEqual(before, removal.Module.Code);
        }
    }

    public sealed partial class VbeCodeNavigationCoverageTests
    {
        [TestMethod]
        public void FindHandlesAggregateLineReadbackFallbackAndTerminalNewline()
        {
            foreach (bool terminal in new[] { false, true })
            {
                var host = new FakeVbe(); var project = new FakeProject { Name = "Projet", Mode = 2 };
                var module = new FakeModule("alpha\r\nalpha"); module.AppendTerminator = terminal; module.TruncateAggregate = !terminal;
                project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = module }); host.VBProjects.Add(project);
                dynamic result = new VbeCodeNavigation(host, new VbeForms(host)).Find(new Request { Project = "Projet", Query = "alpha" });
                Assert.AreEqual(2, (int)result.Matches.Count);
                Assert.AreEqual(2, (int)result.Matches[1].StartLine);
            }
        }

        [TestMethod]
        public void WildcardsTruncateAndWordBoundariesCoverEveryNeighbour()
        {
            var navigation = CreateNavigation(string.Join("\r\n", System.Linq.Enumerable.Repeat("alpha", 205)), "alpha");
            dynamic result = navigation.Find(new Request { Project = "Projet", Query = "a?pha", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(200, (int)result.Matches.Count); Assert.IsTrue((bool)result.Truncated);
            foreach (bool pattern in new[] { false, true })
            {
                navigation = CreateNavigation("alpha alpha! xalpha alphax _alpha alpha_ 0alpha alpha0", "");
                dynamic words = navigation.Find(new Request { Project = "Projet", Query = "alpha", PatternSearch = pattern, WholeWord = true });
                Assert.AreEqual(2, (int)words.Matches.Count);
            }
        }

        [TestMethod]
        public void SelectProcedureValidatesAllInputsAndSelectsTheActualBodyLine()
        {
            const string source = "Sub Run()\r\nEnd Sub";
            var project = new VbeDebugTests.FakeProject { Name = "Projet", Mode = 2 };
            var component = new VbeDebugTests.FakeComponent { Name = "Module1", Type = 1 };
            var module = new VbeDebugTests.FakeModule(component, source); component.CodeModule = module;
            project.VBComponents.Add(component);
            var host = new VbeDebugTests.FakeVbe { ActiveVBProject = project, ActiveCodePane = module.CodePane };
            host.VBProjects.Add(project);
            var nav = new VbeCodeNavigation(host, new VbeForms(host)); var debug = new VbeDebug(host);
            for (int fault = 0; fault < 4; fault++)
            {
                var r = new Request { Project = "Projet", Module = "Module1", Procedure = "Run", ExpectedSha256 = Hash(source) };
                if (fault == 0) r.Procedure = " ";
                if (fault == 1) r.ProcKind = -1;
                if (fault == 2) r.ProcKind = 4;
                if (fault == 3) r.ExpectedSha256 = null;
                Assert.ThrowsException<ArgumentException>(() => nav.SelectProcedure(r, debug));
                Assert.AreEqual(0, module.CodePane.ShowCount);
            }
            var request = new Request { Project = "Projet", Module = "Module1", Procedure = "Run", ExpectedSha256 = Hash(source), ExpectedMode = 2 };
            nav.SelectProcedure(request, debug);
            Assert.AreEqual(1, request.StartLine); Assert.AreEqual(1, module.CodePane.StartLine); Assert.AreEqual(1, module.CodePane.ShowCount);
        }
    }

    public sealed partial class CodeFileEncodingTests
    {
        private static string Decode(byte[] bytes, string encoding, Func<int, Encoding> legacy = null)
        {
            var navigation = new VbeCodeNavigation(null, null);
            if (legacy != null) navigation.LegacySourceEncoding = legacy;
            var method = typeof(VbeCodeNavigation).GetMethod("DecodeCodeFile", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            try { return (string)method.Invoke(navigation, new object[] { bytes, encoding, null }); }
            catch (System.Reflection.TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }

        [TestMethod]
        public void DecodeCoversBomVariantsExplicitCodePagesAndInvalidSequences()
        {
            Assert.ThrowsException<ArgumentException>(() => Decode(new byte[] { 65 }, "unknown"));
            foreach (var codec in new Encoding[] { new UTF8Encoding(true, true), new UnicodeEncoding(false, true, true), new UnicodeEncoding(true, true, true) })
            {
                string name = codec.CodePage == 65001 ? "utf-8" : codec.CodePage == 1200 ? "utf-16le" : "utf-16be";
                byte[] bytes = Combine(codec.GetPreamble(), codec.GetBytes("'été"));
                Assert.AreEqual("'été", Decode(bytes, null));
                Assert.AreEqual("'été", Decode(bytes, " " + name.ToUpperInvariant() + " "));
                Assert.ThrowsException<ArgumentException>(() => Decode(bytes, "windows-1252"));
                Assert.AreEqual("'été", Decode(codec.GetBytes("'été"), name));
            }
            foreach (byte[] bytes in new[] { new byte[] { 255, 254, 0, 0 }, new byte[] { 0, 0, 254, 255 } })
                Assert.ThrowsException<ArgumentException>(() => Decode(bytes, null));
            foreach (string encoding in new[] { "utf-8", "utf-16le", "utf-16be" })
                Assert.ThrowsException<ArgumentException>(() => Decode(new byte[] { 255 }, encoding));
            Assert.ThrowsException<ArgumentException>(() => Decode(new byte[] { 65, 233 }, null));
            Assert.AreEqual("'été", Decode(Encoding.GetEncoding(1252).GetBytes("'été"), "windows-1252"));
            Assert.AreEqual("ASCII", Decode(Encoding.Default.GetBytes("ASCII"), "system-ansi"));
            foreach (byte[] bytes in new[] { new byte[] { 0 }, new byte[] { 254, 0 }, new byte[] { 239, 0, 0 }, new byte[] { 0, 0, 254, 0 } })
                Assert.AreEqual(Encoding.ASCII.GetString(bytes), Decode(bytes, "windows-1252").Replace('þ', '?').Replace('ï', '?'));
        }

        [TestMethod]
        public void InspectionRecognizesBothUtf32ByteOrdersAndBinaryAmbiguity()
        {
            string path = Path.Combine(Path.GetTempPath(), "VBAi-inspect-" + Guid.NewGuid().ToString("N") + ".bas");
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(null));
                foreach (byte[] bytes in new[] { new byte[] { 255, 254, 0, 0 }, new byte[] { 0, 0, 254, 255 } })
                {
                    File.WriteAllBytes(path, bytes); dynamic info = reader.InspectCodeFile(path);
                    Assert.AreEqual("utf-32", (string)info.Bom); Assert.IsNull((object)info.StrictUtf8Valid);
                }
                File.WriteAllBytes(path, new byte[] { 65, 0 }); dynamic binary = reader.InspectCodeFile(path);
                Assert.IsTrue((bool)binary.ContainsNulByte); Assert.IsTrue((bool)binary.ExplicitEncodingRequired); Assert.IsNull((object)binary.DefaultEncoding);
                File.WriteAllBytes(path, new byte[] { 255, 254, 0, 1 });
                Assert.AreEqual("utf-16le", (string)((dynamic)reader.InspectCodeFile(path)).Bom);
                foreach (byte[] prefix in new[] { new byte[] { 255, 254, 0, 0 }, new byte[] { 0, 0, 254, 255 }, new byte[] { 239, 187, 191, 65 }, new byte[] { 254, 255, 0, 65 } })
                    for (int position = 0; position < prefix.Length; position++)
                    {
                        byte[] bytes = (byte[])prefix.Clone(); bytes[position] ^= 1;
                        File.WriteAllBytes(path, bytes); dynamic info = reader.InspectCodeFile(path);
                        Assert.AreEqual(4, (int)info.ByteCount); Assert.AreEqual(64, ((string)info.Sha256).Length);
                    }
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void StrictRoundtripRejectsNonCanonicalLegacyCodecBytes()
        {
            var iso2022 = Encoding.GetEncoding(50220, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            StringAssert.Contains(Assert.ThrowsException<ArgumentException>(() => Decode(new byte[] { 27, 40, 66, 65 }, "system-ansi", _ => iso2022)).Message, "reproduce");
            StringAssert.Contains(Assert.ThrowsException<ArgumentException>(() => Decode(new byte[] { 27, 36, 64, 36, 34, 27, 40, 66 }, "system-ansi", _ => iso2022)).Message, "reproduce");
            var utf8 = new UTF8Encoding(false, true);
            Assert.AreEqual("é", Decode(new byte[] { 195, 169 }, "system-ansi", _ => utf8));
        }
    }

    public sealed partial class VbeProcedureMutationTests
    {
        [TestMethod]
        public void ControlCountingTraversesContainersWithoutTreatingTheirNamesAsControls()
        {
            var nodes = new object[]
            {
                new ControlNode { Kind = "Page", Name = "Target", Children = new object[] { new ControlNode { Kind = "Control", Name = "target" } } },
                new ControlNode { Kind = "Control", Name = "Other" }
            };
            var method = typeof(VbeCodeNavigation).GetMethod("CountControls", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            Assert.AreEqual(1, (int)method.Invoke(null, new object[] { nodes, "TARGET" }));
            Assert.AreEqual(0, (int)method.Invoke(null, new object[] { nodes, "absent" }));
        }
    }

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class CodeFileEncodingTests
    {
        [TestMethod]
        public void InspectionDistinguishesAsciiUtf8BomAndAmbiguousAnsi()
        {
            var root = Path.Combine(Path.GetTempPath(), "VBAi-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                var ascii = Path.Combine(root, "plain.bas");
                File.WriteAllBytes(ascii, Encoding.ASCII.GetBytes("Option Explicit\r\n"));
                dynamic plain = reader.InspectCodeFile(ascii);
                Assert.AreEqual("utf-8", (string)plain.DefaultEncoding);
                Assert.IsFalse((bool)plain.ExplicitEncodingRequired);
                Assert.IsTrue((bool)plain.StrictUtf8Valid);
                Assert.IsFalse((bool)plain.ContentIncluded);
                var utf8 = Path.Combine(root, "utf8.bas");
                File.WriteAllBytes(utf8, Combine(new byte[] { 0xEF, 0xBB, 0xBF }, Encoding.UTF8.GetBytes("' été\r\n")));
                dynamic withBom = reader.InspectCodeFile(utf8);
                Assert.AreEqual("utf-8", (string)withBom.Bom);
                Assert.AreEqual("utf-8", (string)withBom.DefaultEncoding);
                Assert.IsTrue((bool)withBom.ContainsNonAscii);
                Assert.IsFalse((bool)withBom.ExplicitEncodingRequired);
                var ansi = Path.Combine(root, "ansi.bas");
                File.WriteAllBytes(ansi, new byte[] { 0x27, 0x20, 0xE9, 0x0D, 0x0A });
                dynamic ambiguous = reader.InspectCodeFile(ansi);
                Assert.IsNull((object)ambiguous.Bom);
                Assert.IsNull((object)ambiguous.DefaultEncoding);
                Assert.IsFalse((bool)ambiguous.StrictUtf8Valid);
                Assert.IsTrue((bool)ambiguous.ExplicitEncodingRequired);
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [TestMethod]
        public void InspectionReportsUtf16AndRejectsUnsafeFileInputs()
        {
            var root = Path.Combine(Path.GetTempPath(), "VBAi-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile("relative.bas"));
                Assert.ThrowsException<FileNotFoundException>(() => reader.InspectCodeFile(Path.Combine(root, "missing.bas")));
                var empty = Path.Combine(root, "empty.bas");
                File.WriteAllBytes(empty, new byte[0]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(empty));
                var oversized = Path.Combine(root, "oversized.bas");
                File.WriteAllBytes(oversized, new byte[256 * 1024 + 1]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(oversized));
                var utf16 = Path.Combine(root, "utf16.bas");
                File.WriteAllBytes(utf16, Combine(new byte[] { 0xFF, 0xFE }, Encoding.Unicode.GetBytes("' été")));
                dynamic unicode = reader.InspectCodeFile(utf16);
                Assert.AreEqual("utf-16le", (string)unicode.Bom);
                Assert.IsTrue((bool)unicode.ContainsNulByte);
                Assert.IsNull((object)unicode.StrictUtf8Valid);
                Assert.IsFalse((bool)unicode.ExplicitEncodingRequired);
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [TestMethod]
        public void InsertionRequiresExplicitEncodingAndRollsBackCorruptedUnicode()
        {
            var root = Path.Combine(Path.GetTempPath(), "VBAi-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var source = Path.Combine(root, "accent.bas");
                File.WriteAllBytes(source, Encoding.GetEncoding(1252).GetBytes("' été"));
                var module = new VbeSessionContractTests.FakeModule();
                var project = new VbeSessionContractTests.FakeProject
                {
                    Name = "Projet",
                    Mode = 2
                };
                project.VBComponents.Add(new VbeSessionContractTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
                var host = new VbeSessionContractTests.FakeVbe();
                host.VBProjects.Add(project);
                var navigation = new VbeCodeNavigation(host, new VbeForms(host));
                var request = new Request
                {
                    Project = "Projet",
                    Module = "Module1",
                    Path = source,
                    StartLine = 1,
                    ExpectedSha256 = Hash(string.Empty)
                };
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "utf-8";
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "windows-1252";
                module.CorruptNonAsciiOnInsert = true;
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
                Assert.AreEqual(0, module.CountOfLines, "Unicode corruption must roll back all inserted lines.");
                module.CorruptNonAsciiOnInsert = false;
                dynamic inserted = navigation.InsertCodeFile(request);
                Assert.AreEqual("windows-1252", (string)inserted.SourceEncoding);
                Assert.AreEqual(1, (int)inserted.InsertedLineCount);
                StringAssert.Contains(module.Code, "été");
                Assert.IsFalse((bool)inserted.CompilationVerified);
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeCodeNavigationCoverageTests
    {
        [TestMethod]
        public void FindSearchesEveryModuleAndHonorsCaseAndWholeWord()
        {
            var navigation = CreateNavigation("alpha Alpha alphabet\r\nALPHA", "Alpha");
            dynamic all = navigation.Find(new Request { Project = "Projet", Query = "alpha" });
            Assert.AreEqual(5, (int)all.Matches.Count);
            Assert.AreEqual(2, (int)all.SourceVersions.Count);
            Assert.AreEqual("Module1", (string)all.Matches[0].Module);
            Assert.AreEqual(1, (int)all.Matches[0].StartColumn);
            Assert.AreEqual(5, (int)all.Matches[0].EndColumn);
            dynamic words = navigation.Find(new Request { Project = "Projet", Query = "alpha", WholeWord = true });
            Assert.AreEqual(4, (int)words.Matches.Count);
            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "module1", Query = "Alpha", MatchCase = true, WholeWord = true });
            Assert.AreEqual(1, (int)exact.Matches.Count);
            Assert.AreEqual(7, (int)exact.Matches[0].StartColumn);
        }

        [TestMethod]
        public void FindWildcardSearchRespectsWholeWordAndEscapedLiterals()
        {
            var navigation = CreateNavigation("foo_1 foo-2 xfoo_3\r\n[a] [b]", "unused");
            dynamic pattern = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "foo?1", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(1, (int)pattern.Matches.Count);
            dynamic embedded = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "foo?3", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(0, (int)embedded.Matches.Count);
            dynamic literal = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "[?]", PatternSearch = true });
            Assert.AreEqual(2, (int)literal.Matches.Count);
            Assert.AreEqual(2, (int)literal.Matches[0].StartLine);
        }

        [TestMethod]
        public void FindRejectsInvalidQueriesAndUnknownModule()
        {
            var navigation = CreateNavigation("abc", "def");
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = "" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = new string ('x', 201) }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = "***", PatternSearch = true }));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.Find(new Request { Project = "Projet", Module = "Missing", Query = "abc" }));
        }

        [TestMethod]
        public void FindTruncatesLargeResultSetsAtTwoHundredMatches()
        {
            var navigation = CreateNavigation(string.Join("\r\n", Enumerable.Repeat("hit", 220)), "");
            dynamic result = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "hit" });
            Assert.IsTrue((bool)result.Truncated);
            Assert.AreEqual(200, (int)result.Matches.Count);
            Assert.AreEqual(200, (int)result.Matches[199].StartLine);
        }

        [TestMethod]
        public void FindReportsOverlappingMatchesAndTreatsUnicodeAndUnderscoreAsWordCharacters()
        {
            var navigation = CreateNavigation("aaaa\r\néalpha alpha_ alpha é alpha", "");
            dynamic overlapping = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "aa", MatchCase = true });
            Assert.AreEqual(3, (int)overlapping.Matches.Count);
            Assert.AreEqual(1, (int)overlapping.Matches[0].StartColumn);
            Assert.AreEqual(2, (int)overlapping.Matches[1].StartColumn);
            Assert.AreEqual(3, (int)overlapping.Matches[2].StartColumn);
            dynamic words = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "alpha", WholeWord = true });
            Assert.AreEqual(2, (int)words.Matches.Count);
            Assert.AreEqual(2, (int)words.Matches[0].StartLine);
            Assert.AreEqual(15, (int)words.Matches[0].StartColumn);
            Assert.AreEqual(23, (int)words.Matches[1].StartColumn);
        }

        [TestMethod]
        public void FindIncludesEmptyModulesInSourceVersionsAndHonorsWildcardCase()
        {
            var navigation = CreateNavigation("", "Alpha alpha ALPHA");
            dynamic empty = navigation.Find(new Request { Project = "Projet", Query = "missing" });
            Assert.AreEqual(0, (int)empty.Matches.Count);
            Assert.AreEqual(2, (int)empty.SourceVersions.Count);
            Assert.IsFalse((bool)empty.Truncated);
            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "Class1", Query = "A?pha", PatternSearch = true, MatchCase = true });
            Assert.AreEqual(1, (int)exact.Matches.Count);
            Assert.AreEqual(1, (int)exact.Matches[0].StartColumn);
            dynamic folded = navigation.Find(new Request { Project = "Projet", Module = "Class1", Query = "A?pha", PatternSearch = true, MatchCase = false });
            Assert.AreEqual(3, (int)folded.Matches.Count);
            Assert.AreEqual(13, (int)folded.Matches[2].StartColumn);
        }

        [TestMethod]
        public void ProcedureMutationRejectsBreakModeAndWrongComponentWithoutChangingCode()
        {
            var host = new VbeProcedureMutationTests.FakeVbe();
            var project = new VbeProcedureMutationTests.FakeProject
            {
                Name = "Projet",
                Mode = 1
            };
            var component = new VbeProcedureMutationTests.FakeComponent
            {
                Name = "Module1",
                Type = 1
            };
            var module = new VbeProcedureMutationTests.FakeModule(component, "Option Explicit");
            component.CodeModule = module;
            project.VBComponents.Add(component);
            host.VBProjects.Add(project);
            var navigation = new VbeCodeNavigation(host, new VbeForms(host));
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                Text = "Sub Run()\nEnd Sub",
                ExpectedSha256 = Hash(module.Code)
            };
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(request));
            Assert.AreEqual("Option Explicit", module.Code);
            project.Mode = 2;
            component.Type = 3;
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(request));
            Assert.AreEqual("Option Explicit", module.Code);
        }

        [TestMethod]
        public void SelectProcedureRejectsStaleHashBeforeLookingUpProcedureOrSelectingCode()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                ExpectedSha256 = Hash("stale")
            };
            var error = Assert.ThrowsException<InvalidOperationException>(() => navigation.SelectProcedure(request, null));
            StringAssert.Contains(error.Message, "module changed since it was read");
            Assert.AreEqual(0, request.StartLine);
        }

        [TestMethod]
        public void ProcedureCommandsRejectMalformedOrStaleInputBeforeEditing()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                Text = "Sub Run()\nEnd Sub",
                ExpectedSha256 = Hash("Option Explicit")
            };
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(new Request()));
            request.Text = "Sub Run()\nSub Other()\nEnd Sub";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(request));
            request.Text = "Sub Run()\nEnd Sub";
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<ArgumentException>(() => navigation.RemoveProcedure(new Request { Project = "Projet", Module = "Module1", Procedure = "bad name", ExpectedSha256 = request.ExpectedSha256 }));
            Assert.ThrowsException<ArgumentException>(() => navigation.SelectProcedure(new Request { Project = "Projet", Module = "Module1", Procedure = "Run" }, null));
        }

        [TestMethod]
        public void ProcedureValidationCoversEventAndPropertyDeclarations()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var currentHash = Hash("Option Explicit");
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request { EventName = "Click", ObjectName = "Bad Name", ExpectedSha256 = currentHash, ExpectedTreeVersion = "tree" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request { EventName = "Click!", ObjectName = "Button1", ExpectedSha256 = currentHash, ExpectedTreeVersion = "tree" }));
            var property = new Request
            {
                Project = "Projet",
                Module = "Class1",
                Procedure = "Value",
                ProcKind = 3,
                ExpectedSha256 = Hash(""),
                Text = "Public Property Get Value() As String\nEnd Property"
            };
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(property));
            property.Text = "Public Property Let Value(value As String)\nEnd Property";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(property));
            property.ProcKind = 1;
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using System.Text.RegularExpressions;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProcedureMutationTests
    {
        [TestMethod]
        public void ProjectSymbolsReturnsLiveDefinitionsAndPaginates()
        {
            var fixture = new Fixture("Option Explicit");
            fixture.Navigation.CreateProcedure(fixture.Request("Sub Run()\nDebug.Print 1\nEnd Sub"));
            dynamic all = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Limit = 1 });
            Assert.AreEqual(2, (int)all.Total);
            Assert.IsTrue((bool)all.HasMore);
            Assert.AreEqual(0, (int)all.Errors.Count);
            dynamic found = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "run", WholeWord = true });
            Assert.AreEqual(1, (int)found.Total);
            Assert.AreEqual("Run", (string)found.Symbols[0].Name);
            Assert.AreEqual(3, (int)found.Symbols[0].Line);
            Assert.IsFalse(string.IsNullOrEmpty((string)found.Symbols[0].Sha256));
            dynamic exactCase = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "run", WholeWord = true, MatchCase = true });
            Assert.AreEqual(0, (int)exactCase.Total);
        }

        [TestMethod]
        public void CreateReplaceListAndRemoveProcedureRoundTrip()
        {
            var fixture = new Fixture("Option Explicit");
            var create = fixture.Request("Sub Run()\nDebug.Print 1\nEnd Sub");
            dynamic made = fixture.Navigation.CreateProcedure(create);
            Assert.AreEqual(3, (int)made.BodyLine);
            Assert.AreEqual("Run", (string)made.Procedure);
            Assert.IsFalse((bool)made.CompilationVerified);
            dynamic listed = fixture.Navigation.Procedures("Projet", "Module1");
            Assert.AreEqual(1, (int)listed.Procedures.Count);
            Assert.AreEqual("Run", (string)listed.Procedures[0].Name);
            create.ExpectedSha256 = (string)made.Sha256;
            create.Text = "Sub Run()\nDebug.Print 2\nEnd Sub";
            dynamic replaced = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsTrue((bool)replaced.Changed);
            StringAssert.Contains(fixture.Module.Code, "Debug.Print 2");
            create.ExpectedSha256 = (string)replaced.Sha256;
            dynamic unchanged = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsFalse((bool)unchanged.Changed);
            dynamic removed = fixture.Navigation.RemoveProcedure(create);
            Assert.AreEqual(3, (int)removed.RemovedLineCount);
            Assert.IsFalse(fixture.Module.Code.Contains("Run()"));
            Assert.AreEqual(0, (int)((dynamic)fixture.Navigation.Procedures("Projet", "Module1")).Procedures.Count);
        }

        [TestMethod]
        public void FunctionReplacementAndRemovalUseEndFunctionBoundary()
        {
            var fixture = new Fixture("Option Explicit");
            var request = fixture.Request("Public Function Run() As Long\nRun = 1\nEnd Function");
            dynamic created = fixture.Navigation.CreateProcedure(request);
            Assert.AreEqual("Run", (string)created.Procedure);
            request.ExpectedSha256 = (string)created.Sha256;
            request.Text = "Public Function Run() As Long\nRun = 2\nEnd Function";
            dynamic replaced = fixture.Navigation.ReplaceProcedure(request);
            Assert.IsTrue((bool)replaced.Changed);
            StringAssert.Contains(fixture.Module.Code, "Run = 2");
            request.ExpectedSha256 = (string)replaced.Sha256;
            dynamic removed = fixture.Navigation.RemoveProcedure(request);
            Assert.AreEqual(3, (int)removed.RemovedLineCount);
            Assert.IsFalse(fixture.Module.Code.Contains("Function Run"));
        }

        [TestMethod]
        public void StaleVersionAndDuplicateProcedureNeverChangeCode()
        {
            var fixture = new Fixture("Option Explicit");
            var request = fixture.Request("Sub Run()\nEnd Sub");
            request.ExpectedSha256 = Hash("stale");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual("Option Explicit", fixture.Module.Code);
            request.ExpectedSha256 = Hash(fixture.Module.Code);
            fixture.Navigation.CreateProcedure(request);
            var before = fixture.Module.Code;
            request.ExpectedSha256 = Hash(before);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedReplacementRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nDebug.Print 1\r\nEnd Sub");
            var request = fixture.Request("Sub Run()\nDebug.Print 2\nEnd Sub");
            var before = fixture.Module.Code;
            fixture.Module.FailNextInsert = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.ReplaceProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedRemovalRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nEnd Sub");
            var request = fixture.Request(null);
            var before = fixture.Module.Code;
            fixture.Module.PretendStillPresentAfterDelete = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.RemoveProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void EventStubUsesCurrentFormTreeAndCurrentCodeVersion()
        {
            var fixture = new Fixture("", 3);
            fixture.Component.Designer.Controls.AddExisting("Button1");
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request
            {
                Project = "Projet",
                Form = "Module1",
                ObjectName = "Button1",
                EventName = "Click",
                ExpectedTreeVersion = tree.TreeVersion,
                ExpectedSha256 = Hash("")
            };
            dynamic result = fixture.Navigation.CreateEventProcedure(request);
            Assert.AreEqual("Button1_Click", (string)result.Procedure);
            StringAssert.Contains(fixture.Module.Code, "Private Sub Button1_Click()");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
        }

        [TestMethod]
        public void EventStubRejectsStaleTreeAndUnknownControlBeforeEditingCode()
        {
            var fixture = new Fixture("", 3);
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request
            {
                Project = "Projet",
                Form = "Module1",
                ObjectName = "Missing",
                EventName = "Click",
                ExpectedTreeVersion = tree.TreeVersion,
                ExpectedSha256 = Hash("")
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
            request.ObjectName = "UserForm";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
        }
    }
}
