using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbaGitSnapshotTests
    {
        [TestMethod]
        public void SingleUnambiguousFormBlobProducesAnExactFontPlanWithoutChangingSnapshotFiles()
        {
            var component = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            byte[] resource = FormStreamPaddingTests.ContainerResourceBefore(); byte[] original = (byte[])resource.Clone();
            var snapshot = FontSnapshot(component, "OleObjectBlob = \"Form1.frx\":0000\n", resource);
            var bindings = snapshot.FormFonts(component);
            Assert.AreEqual(2, bindings.Length);
            Assert.AreEqual(82700u, BitConverter.ToUInt32(bindings.Single(binding => binding.OwnerPath == "Controls/QualificationExtra").Descriptor, 6));
            CollectionAssert.AreEqual(original, snapshot.Files["Form1.frx"]);
        }

        [TestMethod]
        public void OpaqueDuplicateOrMissingBlobDoesNotInferAPartialFontPlan()
        {
            var component = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
            foreach (string declarations in new[] {
                "", "OleObjectBlob = \"Form1.frx\":0000\nPicture = \"Form1.frx\":0001\n",
                "OleObjectBlob = \"Form1.frx\":0000\nOleObjectBlob = \"Form1.frx\":0000\n" })
                Assert.IsNull(FontSnapshot(component, declarations, resource).FormFonts(component));
            component.HasResources = false;
            Assert.IsNull(FontSnapshot(component, "", null).FormFonts(component));
            component.Type = 2;
            Assert.IsNull(FontSnapshot(component, "", null).FormFonts(component));
        }

        private static VbaGitSnapshot FontSnapshot(VbaGitComponent component, string metadata, byte[] resource)
        {
            var files = new Dictionary<string, byte[]> { [component.FileName] = VbaGitSnapshot.Utf8.GetBytes(metadata + "Attribute VB_Name = \"Form1\"\nOption Explicit\n") };
            if (component.HasResources) files.Add("Form1.frx", resource);
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { component } }, files);
        }
    }
}
