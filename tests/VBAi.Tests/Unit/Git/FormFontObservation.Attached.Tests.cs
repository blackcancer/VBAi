using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormFontObservationTests
    {
        [TestMethod]
        public void AttachedDisabledReturnsWithoutSnapshotPathOrNativeAccess()
        {
            Assert.IsNull(FormFontObservation.Attached.TryBegin(null, null, null, null));
            Assert.IsNull(FormFontObservation.Attached.TryBegin("", null, null, null));
        }

        [TestMethod]
        public void AttachedManifestRequiresFrozenCandidateSourceResourceAndDeclaredOwnerPaths()
        {
            var snapshot = Target();
            foreach (string stage in new[] { "baseline-only", "font-delivery-returned" })
            {
                var value = AttachedValid(snapshot, stage);
                FormFontObservation.ValidateAttachedManifest(value, value.ProjectPath, snapshot, snapshot);
                Assert.IsFalse(Directory.Exists(value.OutputRoot));
                foreach (string field in new[] { "stage", "getter", "mvid", "source", "resource", "project", "nonce", "owners", "output" })
                {
                    var bad = AttachedValid(snapshot, stage);
                    if (field == "stage") bad.Stage = "retry";
                    if (field == "getter") bad.Getter = "Clone";
                    if (field == "mvid") bad.CandidateMvid = Guid.NewGuid().ToString();
                    if (field == "source") bad.FormSha256 = new string('0', 64);
                    if (field == "resource") bad.ResourceSha256 = new string('0', 64);
                    if (field == "project") bad.ProjectPath = Path.Combine(Path.GetTempPath(), "other.xlsm");
                    if (field == "nonce") bad.Nonce = "old";
                    if (field == "owners") bad.OwnerPaths = new[] { "", "Controls/Frame1/Controls/Other" };
                    if (field == "output") bad.OutputRoot += "-reused";
                    Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.ValidateAttachedManifest(
                        bad, value.ProjectPath, snapshot, snapshot), field);
                }
            }
        }

        [TestMethod]
        public void AttachedBaselinePinsExpectedSnapshotAndDeliveryPinsTarget()
        {
            var before = Target(); var changed = WithResource(before, DifferentResource());
            var baseline = AttachedValid(before, "baseline-only");
            FormFontObservation.ValidateAttachedManifest(baseline, baseline.ProjectPath, before, changed);
            var delivery = AttachedValid(before, "font-delivery-returned");
            FormFontObservation.ValidateAttachedManifest(delivery, delivery.ProjectPath, changed, before);
            Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.ValidateAttachedManifest(
                baseline, baseline.ProjectPath, changed, before));
            Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.ValidateAttachedManifest(
                delivery, delivery.ProjectPath, before, changed));
        }

        [TestMethod]
        public void AttachedUniqueGetterHasDurableIntentBeforeExactlyOneGetAndSave()
        {
            var order = new List<string>();
            FormFontObservation.RunAttachedGetter(() => { order.Add("getter"); return 8.27m; },
                () => { order.Add("save"); return new byte[] { 1 }; }, (phase, data) => order.Add(phase));
            CollectionAssert.AreEqual(new[] { "single-getter-intent", "getter", "single-getter-returned",
                "persist-after-intent", "save", "persist-after" }, order.ToArray());
        }

        [TestMethod]
        public void AttachedWriterGetterAndSaveFailuresAreNeverRetried()
        {
            foreach (string fault in new[] { "single-getter-intent", "getter", "single-getter-returned", "persist-after-intent", "save", "persist-after" })
            {
                int gets = 0, saves = 0; var expected = new IOException(fault);
                try
                {
                    FormFontObservation.RunAttachedGetter(() => { gets++; if (fault == "getter") throw expected; return 8.27m; },
                        () => { saves++; if (fault == "save") throw expected; return new byte[] { 1 }; },
                        (phase, data) => { if (fault == phase) throw expected; });
                    Assert.Fail("The exact diagnostic failure must escape.");
                }
                catch (IOException actual) { Assert.AreSame(expected, actual); }
                Assert.AreEqual(fault == "single-getter-intent" ? 0 : 1, gets);
                Assert.AreEqual(fault == "save" || fault == "persist-after" ? 1 : 0, saves);
            }
        }

        [TestMethod]
        public void AttachedGuardRunsAfterIntentImmediatelyBeforeGetterAndSave()
        {
            var order = new List<string>();
            FormFontObservation.RunAttachedGetter(() => { order.Add("getter"); return 8.27m; },
                () => { order.Add("save"); return new byte[] { 1 }; }, (phase, data) => order.Add(phase), () => order.Add("guard"));
            CollectionAssert.AreEqual(new[] { "single-getter-intent", "guard", "getter", "single-getter-returned",
                "persist-after-intent", "guard", "save", "persist-after" }, order.ToArray());
            int gets = 0, saves = 0;
            Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.RunAttachedGetter(
                () => { gets++; return 0; }, () => { saves++; return new byte[] { 1 }; }, (phase, data) => { },
                () => { throw new InvalidOperationException("owner changed"); }));
            Assert.AreEqual(0, gets); Assert.AreEqual(0, saves);
        }

        [TestMethod]
        public void AttachedSchemaRejectsExtraFieldsWrongCloneTypeAndOversizedBytes()
        {
            var value = AttachedValid(Target(), "baseline-only");
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            string text = json.Serialize(value);
            Assert.AreEqual(value.Getter, FormFontObservation.ParseAttachedManifest(System.Text.Encoding.UTF8.GetBytes(text)).Getter);
            foreach (string invalid in new[] { text.TrimEnd('}') + ",\"Replay\":true}",
                text.Replace("\"DetachedFrameClone\":false", "\"DetachedFrameClone\":\"true\""),
                text.Replace("\"ResourceSha256\"", "\"resourceSha256\"") })
                Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.ParseAttachedManifest(System.Text.Encoding.UTF8.GetBytes(invalid)));
            Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.ParseAttachedManifest(new byte[16385]));
        }

        [TestMethod]
        public void AttachedOversizedManifestRefusesBeforeParsingOrEvidenceCreation()
        {
            string file = Path.GetTempFileName();
            try
            {
                File.WriteAllBytes(file, new byte[16385]);
                Assert.ThrowsException<InvalidOperationException>(() => FormFontObservation.Attached.TryBegin(file, null, null, null));
            }
            finally { File.Delete(file); }
        }

        [DataTestMethod]
        [DataRow("baseline-only")]
        [DataRow("font-delivery-returned")]
        public void AttachedArmingRequiresFileManifestAndFreshDirectoryWithinFrozenEvidenceRoot(string stage)
        {
            WithAttachedFiles(folder =>
            {
                var snapshot = Target(); var value = AttachedValid(snapshot, stage);
                value.ProjectPath = Path.Combine(folder, "owned.xlsm"); File.WriteAllBytes(value.ProjectPath, new byte[] { 1 });
                value.OutputRoot = Path.Combine(folder, value.Nonce);
                string manifest = Path.Combine(folder, value.Nonce + ".attached-font.json");
                File.WriteAllText(manifest, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(value));
                byte[] frozenManifest = File.ReadAllBytes(manifest);
                var attached = FormFontObservation.Attached.TryBegin(manifest, value.ProjectPath, snapshot, snapshot, folder);
                Assert.IsNotNull(attached); Assert.AreEqual(stage, attached.Stage); Assert.AreEqual("Form1", attached.FormName);
                Assert.IsTrue(Directory.Exists(value.OutputRoot));
                string receipt = Directory.GetFiles(value.OutputRoot).Single();
                Assert.AreEqual("001-armed.json", Path.GetFileName(receipt));
                byte[] frozenReceipt = File.ReadAllBytes(receipt);
                CollectionAssert.AreEqual(frozenManifest, File.ReadAllBytes(manifest));
                Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontObservation.Attached.TryBegin(manifest, value.ProjectPath, snapshot, snapshot, folder));
                CollectionAssert.AreEqual(frozenReceipt, File.ReadAllBytes(receipt), "An existing observation must never be truncated or replayed.");
                Assert.AreEqual(1, Directory.GetFiles(value.OutputRoot).Length);
            });
        }

        [DataTestMethod]
        [DataRow("manifest-outside")]
        [DataRow("output-outside")]
        [DataRow("output-file")]
        public void AttachedArmingRejectsWrongPathScopeAndObjectTypeBeforeCreatingAnyReceipt(string invalid)
        {
            WithAttachedFiles(folder =>
            {
                string evidence = Path.Combine(folder, "evidence"), foreign = Path.Combine(folder, "foreign");
                Directory.CreateDirectory(evidence); Directory.CreateDirectory(foreign);
                var snapshot = Target(); var value = AttachedValid(snapshot, "baseline-only");
                value.ProjectPath = Path.Combine(folder, "owned.xlsm"); File.WriteAllBytes(value.ProjectPath, new byte[] { 1 });
                value.OutputRoot = Path.Combine(invalid == "output-outside" ? foreign : evidence, value.Nonce);
                string manifest = Path.Combine(invalid == "manifest-outside" ? foreign : evidence, value.Nonce + ".attached-font.json");
                if (invalid == "output-file") File.WriteAllBytes(value.OutputRoot, new byte[] { 2 });
                File.WriteAllText(manifest, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(value));
                byte[] frozenManifest = File.ReadAllBytes(manifest);
                if (invalid == "output-file")
                    Assert.ThrowsException<InvalidOperationException>(() =>
                        FormFontObservation.Attached.TryBegin(manifest, value.ProjectPath, snapshot, snapshot, evidence));
                else
                    Assert.ThrowsException<ArgumentException>(() =>
                        FormFontObservation.Attached.TryBegin(manifest, value.ProjectPath, snapshot, snapshot, evidence));
                Assert.IsFalse(Directory.Exists(value.OutputRoot), "No output directory or armed receipt may precede scope validation.");
                CollectionAssert.AreEqual(frozenManifest, File.ReadAllBytes(manifest));
                if (invalid == "output-file") CollectionAssert.AreEqual(new byte[] { 2 }, File.ReadAllBytes(value.OutputRoot));
                Assert.AreEqual(0, Directory.GetFiles(folder, "*-armed.json", SearchOption.AllDirectories).Length);
            });
        }

        private static void WithAttachedFiles(Action<string> action)
        {
            string folder = Path.Combine(Path.GetTempPath(), "VBAi-attached-arming-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string previous = Environment.GetEnvironmentVariable(FormFontObservation.ManifestVariable);
            try { Environment.SetEnvironmentVariable(FormFontObservation.ManifestVariable, null); action(folder); }
            finally { Environment.SetEnvironmentVariable(FormFontObservation.ManifestVariable, previous); Directory.Delete(folder, true); }
        }

        private static FormFontObservation.AttachedManifest AttachedValid(VbaGitSnapshot snapshot, string stage)
        {
            string nonce = Guid.NewGuid().ToString("N");
            string frame = snapshot.FormFonts(snapshot.Manifest.Components.Single()).Single(item => item.Type == 14).OwnerPath;
            return new FormFontObservation.AttachedManifest
            {
                ProjectPath = ProjectPath(), FormName = "Form1", Stage = stage, Getter = "Size",
                CandidateMvid = typeof(FormFontObservation).Module.ModuleVersionId.ToString("D"),
                FormSha256 = Hash(snapshot.Files["Form1.frm"]), ResourceSha256 = Hash(snapshot.Files["Form1.frx"]),
                Nonce = nonce, OutputRoot = Path.Combine(Path.GetTempPath(), "VBAi-attached-font", nonce),
                OwnerPaths = new[] { "", frame }
            };
        }
    }
}
