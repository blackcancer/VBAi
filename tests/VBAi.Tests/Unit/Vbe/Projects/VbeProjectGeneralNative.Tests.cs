using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeProjectGeneralNativeTests
    {
        [DataTestMethod]
        [DataRow(1252, false, "C:\\été-œuvre.chm")]
        [DataRow(65001, false, "C:\\日本-😀.chm")]
        [DataRow(1252, true, "C:\\日本-😀.chm")]
        public void NativeTextRepresentationPreservesValidExactStrings(int codePage, bool unicode, string value)
            => VbeProjectGeneralNative.RequireExactTextRepresentation(value, codePage, unicode);
        [DataTestMethod]
        [DataRow("C:\\日本.chm")]
        [DataRow("C:\\−.chm")]
        [DataRow("C:\\Ａ.chm")]
        public void AnsiHelpFileRefusesUnavailableCharactersAndBestFitSubstitutions(string value)
            => Assert.ThrowsException<VbeProjectGeneralOperation.TextRepresentationRefusedException>(() => VbeProjectGeneralNative.RequireExactTextRepresentation(value, 1252, false));
        [DataTestMethod]
        [DataRow(1252, false)]
        [DataRow(65001, false)]
        [DataRow(1252, true)]
        public void InvalidUtf16NeverBecomesReplacementText(int codePage, bool unicode)
        {
            foreach (string value in new[] { "C:\\" + (char)0xD800 + ".chm", "C:\\" + (char)0xDC00 + ".chm" })
                Assert.ThrowsException<VbeProjectGeneralOperation.TextRepresentationRefusedException>(() => VbeProjectGeneralNative.RequireExactTextRepresentation(value, codePage, unicode));
        }
        [TestMethod]
        public void InvalidNativeCodePageCannotSilentlyUseTheProcessDefault()
            => Assert.ThrowsException<InvalidOperationException>(() => VbeProjectGeneralNative.RequireExactTextRepresentation("C:\\help.chm", 0, false));
        [TestMethod]
        public void UnsupportedCodePageIsNotARepresentationRefusalEligibleForCancel()
            => Assert.ThrowsException<ArgumentOutOfRangeException>(() => VbeProjectGeneralNative.RequireExactTextRepresentation("C:\\help.chm", int.MaxValue, false));
        [DataTestMethod]
        [DataRow(4941)]
        [DataRow(4940)]
        [DataRow(4948)]
        [DataRow(4949)]
        [DataRow(4958)]
        public void ObservedGeneralEditsRequireExactVisibleEnabledWritableContract(int id)
        {
            Assert.IsTrue(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Edit", 0x50010080, true, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id + 1, id, "Edit", 0x50010080, true, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Static", 0x50010080, true, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Edit", 0x500100A0, true, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Edit", 0x50010880, true, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Edit", 0x50010080, false, true));
            Assert.IsFalse(VbeProjectGeneralNative.MatchesEditableContract(id, id, "Edit", 0x50010080, true, false));
        }
        [DataTestMethod]
        [DataRow(1, "OK", true)]
        [DataRow(1, "&OK", true)]
        [DataRow(2, "Annuler", true)]
        [DataRow(2, "&Annuler", true)]
        [DataRow(2, "Cancel", true)]
        [DataRow(2, "&Cancel", true)]
        [DataRow(1, "Cancel", false)]
        [DataRow(2, "OK", false)]
        [DataRow(9, "OK", false)]
        [DataRow(1, "Oui", false)]
        [DataRow(2, "Non", false)]
        [DataRow(1, "O K", false)]
        public void CloseContractNeverSubstitutesOtherButtonsOrUnobservedCaptions(int id, string caption, bool expected)
            => Assert.AreEqual(expected, VbeProjectGeneralNative.MatchesCloseCaption(id, caption));
        [TestMethod]
        public void NativeOptionsVersionIsReopenStableButEveryGeneralFieldInvalidatesIt()
        {
            var before = VbeProjectGeneralOperationTests.Snapshot(); var same = VbeProjectGeneralOperationTests.Copy(before);
            same.Dialog = new IntPtr(900); same.Context = new IntPtr(901);
            Assert.AreEqual(before.OptionsVersion, same.OptionsVersion); Assert.IsFalse(before.SameNative(same));
            foreach (string field in new[] { "name", "description", "helpfile", "context", "compilation" })
            {
                var changed = VbeProjectGeneralOperationTests.Copy(before);
                if (field == "name") changed.Name = "other"; if (field == "description") changed.Description = "other";
                if (field == "helpfile") changed.HelpFile = "C:\\é.chm"; if (field == "context") changed.ContextText = "321"; if (field == "compilation") changed.Compilation = "DEBUG=1";
                Assert.AreNotEqual(before.OptionsVersion, changed.OptionsVersion, field);
            }
        }
        [TestMethod]
        public void LengthPrefixedUnicodeOptionsDoNotAliasAdjacentFields()
        {
            var a = VbeProjectGeneralOperationTests.Snapshot(); var b = VbeProjectGeneralOperationTests.Copy(a);
            a.Description = "ab"; a.HelpFile = "c"; b.Description = "a"; b.HelpFile = "bc";
            Assert.AreNotEqual(a.OptionsVersion, b.OptionsVersion);
            a.HelpFile = "C:\\été.chm"; b.HelpFile = "C:\\ete.chm"; Assert.AreNotEqual(a.OptionsVersion, b.OptionsVersion);
        }
    }
}
