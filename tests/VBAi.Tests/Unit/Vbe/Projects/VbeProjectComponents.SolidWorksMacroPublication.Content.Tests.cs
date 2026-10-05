using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeSolidWorksMacroPublicationTests
    {
        private sealed class NativeDesignerPair : IDisposable
        {
            internal readonly Fixture Owner = new Fixture();
            internal readonly JavaScriptSerializer Json = new JavaScriptSerializer();
            internal readonly Dictionary<string, object> Left = (Dictionary<string, object>)Native09PublicationPair(false)["Designer"];
            internal readonly Dictionary<string, object> Right = (Dictionary<string, object>)Native09PublicationPair(true)["Designer"];
            internal readonly string LeftPath, RightPath;
            internal NativeDesignerPair()
            {
                LeftPath = System.IO.Path.Combine(Owner.Root, "left.frm"); RightPath = System.IO.Path.Combine(Owner.Root, "right.frm");
                Write(LeftPath, ""); Write(RightPath, "");
            }
            internal void Write(string path, string extension)
            {
                string header = Native09ExportHeader.Replace("End\r\n", extension + "End\r\n");
                header = header.Replace("Q020PublishedForm.frx", System.IO.Path.GetFileName(System.IO.Path.ChangeExtension(path, ".frx")));
                File.WriteAllText(path, header + "Attribute VB_Name = \"Q020PublishedForm\"\r\nOption Explicit", System.Text.Encoding.Default);
                File.WriteAllBytes(System.IO.Path.ChangeExtension(path, ".frx"), new byte[] { 1, 2, 3 });
            }
            internal List<Dictionary<string, object>> Properties(Dictionary<string, object> tree) => ((IEnumerable)tree["Properties"]).Cast<Dictionary<string, object>>().ToList();
            internal bool Compare() => VbeProjectComponents.PublicationDesignerContentEquals(
                new VbeProjectComponents.PublicationComponent { Name = "Q020PublishedForm", Type = 3, DesignerJson = Json.Serialize(Left) },
                new VbeProjectComponents.PublicationComponent { Name = "Q020PublishedForm", Type = 3, DesignerJson = Json.Serialize(Right) }, LeftPath, RightPath);
            internal void SetBoth(string name, object value)
            {
                foreach (var tree in new[] { Left, Right })
                {
                    var properties = Properties(tree); var property = properties.SingleOrDefault(x => (string)x["Name"] == name);
                    if (property == null) { property = new Dictionary<string, object> { ["Name"] = name, ["Kind"] = "scalar", ["Error"] = null }; properties.Add(property); }
                    property["Value"] = value; tree["Properties"] = properties.Cast<object>().ToArray();
                }
            }
            public void Dispose() => Owner.Dispose();
        }
        [TestMethod, TestCategory("Unit")]
        public void PublicationDesignerNative09PairUsesOnlyProvedRootDefaultsAndKeepsRawTrees()
        {
            using (var f = new NativeDesignerPair())
            {
                string beforeLeft = f.Json.Serialize(f.Left), beforeRight = f.Json.Serialize(f.Right);
                Assert.IsTrue(f.Compare()); Assert.AreEqual(beforeLeft, f.Json.Serialize(f.Left)); Assert.AreEqual(beforeRight, f.Json.Serialize(f.Right));
                Assert.AreNotEqual(f.Left["TreeVersion"], f.Right["TreeVersion"]);
            }
        }
        [DataTestMethod, TestCategory("Unit"), DataRow("HelpContextID"), DataRow("ShowModal"), DataRow("WhatsThisButton"), DataRow("WhatsThisHelp"), DataRow("Visible")]
        public void PublicationDesignerExplicitNondefaultRequiresBothGettersAndMatchingHeaders(string name)
        {
            using (var f = new NativeDesignerPair())
            {
                object value = name == "HelpContextID" ? (object)42 : name == "ShowModal" || name == "Visible" ? false : true;
                string token = name == "HelpContextID" ? "42" : (bool)value ? "-1" : "0";
                f.SetBoth(name, value); f.Write(f.LeftPath, "   " + name + " = " + token + "\r\n"); f.Write(f.RightPath, "   " + name + " = " + token + "\r\n");
                Assert.IsTrue(f.Compare());
                f.Left["Properties"] = f.Properties(f.Left).Where(p => (string)p["Name"] != name).Cast<object>().ToArray();
                Assert.ThrowsException<InvalidOperationException>(() => f.Compare());
            }
        }
        [DataTestMethod, TestCategory("Unit"), DataRow("malformed"), DataRow("foreign-frx"), DataRow("missing-frx"), DataRow("duplicate-root"), DataRow("unknown-header"), DataRow("read-error")]
        public void PublicationDesignerHeaderAndGetterFailuresRemainStrict(string fault)
        {
            using (var f = new NativeDesignerPair())
            {
                if (fault == "malformed") File.WriteAllText(f.RightPath, "Not a form header");
                if (fault == "foreign-frx") File.WriteAllText(f.RightPath, File.ReadAllText(f.RightPath).Replace("right.frx", "foreign.frx"));
                if (fault == "missing-frx") File.Delete(System.IO.Path.ChangeExtension(f.RightPath, ".frx"));
                if (fault == "duplicate-root") { var properties = f.Properties(f.Right); properties.Add(properties[0]); f.Right["Properties"] = properties.Cast<object>().ToArray(); }
                if (fault == "read-error") f.Properties(f.Right).Single(p => (string)p["Name"] == "CanUndo")["Error"] = "Undo getter unreadable";
                if (fault == "unknown-header") { f.Write(f.RightPath, "   UnknownPersistedSetting = 42\r\n"); Assert.IsFalse(f.Compare()); }
                else Assert.ThrowsException<InvalidOperationException>(() => f.Compare());
            }
        }
        [DataTestMethod, TestCategory("Unit"), DataRow("\r\n", 1), DataRow("\r\n\r\n", 2), DataRow("\n", 0), DataRow(" \r\n", 0), DataRow("'changed\r\n", 0)]
        public void PublicationPrefixReconciliationAcceptsOnlyExactSurplusCrLf(string prefix, int count)
        {
            const string source = "Option Explicit\r\nSub Owned()\r\nEnd Sub";
            Assert.AreEqual(count, VbeProjectComponents.PublicationSurplusImportedCrLfPrefix(source, prefix + source));
            Assert.AreEqual(0, VbeProjectComponents.PublicationSurplusImportedCrLfPrefix(source, prefix + source + "changed"));
        }
    }
}