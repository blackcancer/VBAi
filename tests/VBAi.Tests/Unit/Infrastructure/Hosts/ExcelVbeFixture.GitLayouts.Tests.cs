using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Guards native source construction and the durable font evidence without opening a host.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureGitLayoutsTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NestedLayoutNeverReplacesOrReadsTheInheritedFrameFont(bool qualification)
        {
            string previous = Environment.GetEnvironmentVariable(UserFormQualificationFonts.OptIn);
            try
            {
                Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, qualification ? "1" : null);
                var frame = new FrameProbe();
                typeof(ExcelVbeFixture).GetMethod("PrepareNestedGitLayout", BindingFlags.NonPublic | BindingFlags.Static)
                    .Invoke(null, new object[] { frame });
                Assert.AreEqual(0, frame.FontGets);
                Assert.AreEqual(0, frame.FontSets, "The source font must not pass through a new external StdFont.");
                Assert.AreEqual("Synthetic frame", frame.Caption);
                var multi = (MultiPageProbe)frame.Controls.Added.Single();
                Assert.AreEqual("QualificationMultiPage", multi.Name);
                Assert.AreEqual(220d, multi.Width);
                var page = multi.Pages.Value;
                Assert.AreEqual("Synthetic page", page.Caption);
                var text = page.Controls.Added.Single();
                Assert.AreEqual("QualificationNestedText", text.Name);
                Assert.AreEqual("Original nested text", text.Text);
            }
            finally { Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, previous); }
        }

        [TestMethod]
        public void SerializedFontEvidenceKeepsEveryOwnerTypeAndExactDescriptor()
        {
            var snapshot = Snapshot();
            byte[] original = (byte[])snapshot.Files["Form1.frx"].Clone();
            var bindings = snapshot.FormFonts(snapshot.Manifest.Components[0]);
            Assert.IsTrue(bindings.Length > 1, "The synthetic resource has both root and nested font bindings.");
            var serializer = new JavaScriptSerializer();
            var records = serializer.Deserialize<Dictionary<string, object>[]>(
                serializer.Serialize(ExcelVbeFixture.DescribeGitFontBindings(snapshot, "Form1")));
            Assert.AreEqual(bindings.Length, records.Length);
            for (int i = 0; i < records.Length; i++)
            {
                Assert.AreEqual(3, records[i].Count, "Internal readonly fields must not serialize as an empty object.");
                Assert.AreEqual(bindings[i].OwnerPath, records[i]["OwnerPath"]);
                Assert.AreEqual(bindings[i].Type, Convert.ToUInt32(records[i]["Type"]));
                Assert.AreEqual(BitConverter.ToString(bindings[i].Descriptor).Replace("-", ""), records[i]["DescriptorHex"]);
            }
            CollectionAssert.AreEqual(original, snapshot.Files["Form1.frx"], "Describing native evidence cannot rewrite a resource.");
        }

        [TestMethod]
        public void FontEvidenceRejectsAnAbsentFormInsteadOfReturningAnotherOwner()
        {
            var snapshot = Snapshot();
            Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.DescribeGitFontBindings(snapshot, "Missing"));
        }

        private static VbaGitSnapshot Snapshot()
        {
            var component = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { component } },
                new Dictionary<string, byte[]> {
                    ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("OleObjectBlob = \"Form1.frx\":0000\nAttribute VB_Name = \"Form1\"\nOption Explicit\n"),
                    ["Form1.frx"] = FormStreamPaddingTests.ContainerResourceBefore()
                });
        }

        public sealed class FrameProbe
        {
            public int FontGets, FontSets;
            public string Caption { get; set; }
            public ControlCollectionProbe Controls { get; } = new ControlCollectionProbe();
            public object Font
            {
                get { FontGets++; throw new InvalidOperationException("Source construction read the inherited font."); }
                set { FontSets++; throw new InvalidOperationException("Source construction replaced the inherited font."); }
            }
        }
        public class ControlProbe
        {
            public string Name { get; set; }
            public string Text { get; set; }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }
        public sealed class MultiPageProbe : ControlProbe
        {
            public PagesProbe Pages { get; } = new PagesProbe();
        }
        public sealed class PagesProbe
        {
            public PageProbe Value { get; } = new PageProbe();
            public PageProbe Item(int index) { Assert.AreEqual(0, index); return Value; }
        }
        public sealed class PageProbe
        {
            public string Caption { get; set; }
            public ControlCollectionProbe Controls { get; } = new ControlCollectionProbe();
        }
        public sealed class ControlCollectionProbe
        {
            public readonly List<ControlProbe> Added = new List<ControlProbe>();
            public ControlProbe Add(string kind, string name, bool visible)
            {
                Assert.IsTrue(visible);
                Assert.IsTrue(kind == "Forms.MultiPage.1" || kind == "Forms.TextBox.1");
                ControlProbe control = kind == "Forms.MultiPage.1" ? new MultiPageProbe() : new ControlProbe();
                control.Name = name;
                Added.Add(control);
                return control;
            }
        }
    }
}
