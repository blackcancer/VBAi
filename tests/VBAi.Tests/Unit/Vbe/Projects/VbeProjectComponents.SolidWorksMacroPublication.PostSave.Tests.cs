namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    public sealed partial class VbeSolidWorksMacroPublicationTests
    {
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationVerifiedSaveAllowsDocumentedSavedFlagTransitionWithoutChangingSource()
        {
            using (var f = new Fixture())
            {
                f.Service.PublicationSave = r => {
                    f.SaveAttempts++;
                    foreach (var c in f.Target.VBComponents.Items.Where(c => c.Type != 100)) c.Saved = true;
                    return Task.FromResult<object>(new { Verified = true, Uncertain = false });
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsTrue(result.Verified, result.Error);
                Assert.AreEqual(1, f.SaveAttempts);
                Assert.IsFalse(f.Source.VBComponents.Items[0].Saved);
                Assert.IsNull(result.FinalMismatch);
                Assert.AreEqual(0, System.IO.Directory.GetFiles(result.StagingPath, "final-mismatch-*.json").Length);
                Assert.IsFalse(result.RetryAllowed);
            }
        }

        [DataTestMethod, TestCategory("Unit")]
        [DataRow("code"), DataRow("designer-open"), DataRow("unknown-scalar")]
        public async Task PublicationVerifiedSaveStillRefusesRealComponentDriftWithCapturedAfterSaveDiagnostic(string fault)
        {
            using (var f = new Fixture())
            {
                f.Service.PublicationSave = r => {
                    f.SaveAttempts++;
                    var c = f.Target.VBComponents.Items.Single(x => x.Name == "Module1");
                    c.Saved = true;
                    if (fault == "code") c.CodeModule.Source += "\r\n' changed";
                    if (fault == "designer-open") c.HasOpenDesigner = true;
                    if (fault == "unknown-scalar") c.UnknownScalar = 2;
                    return Task.FromResult<object>(new { Verified = true, Uncertain = false });
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Uncertain); Assert.IsTrue(result.Terminal);
                Assert.AreEqual(1, f.SaveAttempts); Assert.IsNotNull(result.Save);
                Assert.AreEqual("AfterSave", result.FinalMismatch.Phase);
                Assert.AreEqual(1, result.FinalMismatch.ComponentOrdinal);
                Assert.IsTrue(result.FinalMismatch.CanonicalEqual);
                Assert.IsFalse(result.RetryAllowed);
            }
        }
        [DataTestMethod, TestCategory("Unit")]
        [DataRow(1), DataRow(2), DataRow(3)]
        public void PublicationAdditionalObservationPreservesRawComponentVersionCalculation(int type)
        {
            using (var f = new Fixture())
            {
                var component = f.Source.VBComponents.Items[0]; component.Type = type;
                f.Service.PublicationFormTree = (p, n) => new { TreeVersion = "raw-tree", Properties = new object[0], Controls = new object[0] };
                var read = typeof(VbeProjectComponents).GetMethod("ReadPublicationComponent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var observed = (VbeProjectComponents.PublicationComponent)read.Invoke(f.Service, new object[] { f.Source.Name, component });
                string expected;
                if (type == 3)
                {
                    using (var sha = System.Security.Cryptography.SHA256.Create())
                        expected = BitConverter.ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(observed.ComponentSnapshotJson))).Replace("-", "").ToLowerInvariant();
                }
                else expected = (string)((dynamic)f.Service.ComponentProperties(f.Source.Name, component.Name)).Version;
                Assert.AreEqual(expected, observed.Version);
                component.Saved = true;
                var changed = (VbeProjectComponents.PublicationComponent)read.Invoke(f.Service, new object[] { f.Source.Name, component });
                Assert.AreNotEqual(observed.Version, changed.Version, "Raw concurrency must retain the Saved flag.");
                Assert.IsTrue(VbeProjectComponents.PublicationPostSaveMetadataEquals(observed, changed));
            }
        }

        [DataTestMethod, TestCategory("Unit")]
        [DataRow("true-false"), DataRow("saved-error"), DataRow("saved-nonbool"), DataRow("saved-type"), DataRow("unknown-property")]
        public void PublicationPostSaveMetadataAllowanceIsStrictAndOperationSpecific(string fault)
        {
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            Func<bool, object> descriptor = value => new System.Collections.Generic.Dictionary<string, object> {
                ["Name"] = "Saved", ["Kind"] = "scalar", ["Type"] = "System.Boolean", ["Value"] = value, ["Error"] = null, ["ReadOnly"] = false };
            var oldRow = (System.Collections.Generic.Dictionary<string, object>)descriptor(fault == "true-false");
            var newRow = (System.Collections.Generic.Dictionary<string, object>)descriptor(fault != "true-false");
            if (fault == "saved-error") newRow["Error"] = "Native getter failed";
            if (fault == "saved-nonbool") newRow["Value"] = "True";
            if (fault == "saved-type") newRow["Type"] = "System.String";
            var left = new VbeProjectComponents.PublicationComponent { Type = 1, ComponentSnapshotJson = serializer.Serialize(new { Properties = new object[] { oldRow }, Unknown = 1 }) };
            var right = new VbeProjectComponents.PublicationComponent { Type = 1, ComponentSnapshotJson = serializer.Serialize(new { Properties = new object[] { newRow }, Unknown = fault == "unknown-property" ? 2 : 1 }) };
            bool equal;
            try { equal = VbeProjectComponents.PublicationPostSaveMetadataEquals(left, right); }
            catch (InvalidOperationException) { equal = false; }
            Assert.IsFalse(equal);
        }

        [DataTestMethod, TestCategory("Unit"), DataRow("code"), DataRow("canonical")]
        public async Task PublicationPostSaveExportClaimCannotHideComponentCodeDrift(string fault)
        {
            using (var f = new Fixture())
            {
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request(), recordClaim: claim => {
                    if (claim.Phase == "BeforePostSaveVerificationExport") {
                        var current = f.Target.VBComponents.Items.Single(c => c.Name == "Module1");
                        if (fault == "code") current.CodeModule.Source += "\r\n' changed after save";
                        else { var replacement = new Component(current.Name, current.Type); replacement.CodeModule.Source = current.CodeModule.Source; f.Target.VBComponents.Items.Remove(current); f.Target.VBComponents.Items.Add(replacement); }
                    }
                });
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Uncertain); Assert.AreEqual(1, f.SaveAttempts);
                Assert.AreEqual("AfterSave", result.FinalMismatch.Phase); Assert.AreEqual(fault != "code", result.FinalMismatch.CodeEqual);
                Assert.AreEqual(fault != "canonical", result.FinalMismatch.CanonicalEqual);
                Assert.AreEqual(1, result.FinalMismatch.ComponentOrdinal); Assert.IsFalse(result.RetryAllowed);
            }
        }    }
}
