using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using Row = VBAi.VbeNativePaletteState.ColorRow;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativePaletteStateTests
    {
        [TestMethod]
        public void RebasePreservesEveryManualFieldAndRestoresOnlyOwnedAppliedValues()
        {
            foreach (bool startApplied in new[] { false, true })
                foreach (string field in new[] { "Foreground", "Background", "Indicator" })
                    for (int index = 0; index < 10; index++)
                    {
                        var state = NativePaletteFixture.State();
                        var current = (startApplied ? state.Applied : state.Original).Select(r => new Row { Name = r.Name, Foreground = r.Foreground, Background = r.Background, Indicator = r.Indicator }).ToArray();
                        var property = typeof(Row).GetProperty(field);
                        int original = (int)property.GetValue(state.Original[index]);
                        int applied = (int)property.GetValue(state.Applied[index]);
                        int manual = Enumerable.Range(0, 17).First(value => value != original && value != applied);
                        property.SetValue(current[index], manual);
                        var rebased = state.Rebase(current);
                        Assert.AreEqual(manual, property.GetValue(rebased.Original[index]));
                        for (int row = 0; row < 10; row++)
                            foreach (string component in new[] { "Foreground", "Background", "Indicator" })
                            {
                                var value = typeof(Row).GetProperty(component);
                                Assert.AreEqual(row == index && component == field ? manual : (int)value.GetValue(state.Original[row]), value.GetValue(rebased.Original[row]));
                            }
                        rebased.Validate("7.1");
                        Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), state.Original));
                    }
            var invalid = NativePaletteFixture.State();
            Assert.ThrowsException<InvalidDataException>(() => invalid.Rebase(null));
            var renamed = NativePaletteFixture.Rows(); renamed[0].Name = "Different category";
            Assert.ThrowsException<InvalidOperationException>(() => invalid.Rebase(renamed));
        }

        [TestMethod]
        public void AtomicReplacementArchivesOriginalAndPreservesItOnFailure()
        {
            using (var fixture = new NativePaletteFixture())
            {
                var state = NativePaletteFixture.State(); state.SaveNew(fixture.PathName);
                string original = File.ReadAllText(fixture.PathName);
                var current = NativePaletteFixture.Rows(); current[0].Foreground = 12;
                var replacement = state.Rebase(current);
                using (File.Open(fixture.PathName, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    Assert.ThrowsException<IOException>(() => replacement.SaveReplacing(fixture.PathName));
                Assert.AreEqual(original, File.ReadAllText(fixture.PathName));
                Assert.AreEqual(0, Directory.GetFiles(fixture.DirectoryPath, "*.tmp").Length);
                string archive = replacement.SaveReplacing(fixture.PathName);
                Assert.AreEqual(original, File.ReadAllText(archive));
                Assert.AreEqual(12, VbeNativePaletteState.Load(fixture.PathName, "7.1").Original[0].Foreground);
                Assert.AreEqual(0, Directory.GetFiles(fixture.DirectoryPath, "*.tmp").Length);
                Assert.ThrowsException<FileNotFoundException>(() => replacement.SaveReplacing(Path.Combine(fixture.DirectoryPath, "missing.json")));
                Assert.AreEqual(0, Directory.GetFiles(fixture.DirectoryPath, "*.tmp").Length);
                replacement.Applied = null;
                Assert.ThrowsException<InvalidDataException>(() => replacement.SaveReplacing(fixture.PathName));
            }
        }
        [TestMethod]
        public void EveryInvalidRowBoundaryIsRejectedBeforeDarkMapping()
        {
            Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.ValidateRows(null));
            Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.ValidateRows(new Row[9]));
            Action<Row>[] corrupt = { row => row.Name = null, row => row.Name = " ", row => row.Name = "Category 1",
                row => row.Foreground = -1, row => row.Foreground = 17, row => row.Background = -1,
                row => row.Background = 17, row => row.Indicator = -1, row => row.Indicator = 17 };
            foreach (var change in corrupt)
            {
                var rows = NativePaletteFixture.Rows(); change(rows[0]);
                Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.Dark(rows));
            }
            var absent = NativePaletteFixture.Rows(); absent[0] = null;
            Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.ValidateRows(absent));
            var boundaries = NativePaletteFixture.Rows();
            boundaries[0].Foreground = boundaries[0].Background = boundaries[0].Indicator = 0;
            boundaries[1].Foreground = boundaries[1].Background = boundaries[1].Indicator = 16;
            VbeNativePaletteState.ValidateRows(boundaries);
        }

        [TestMethod]
        public void EqualityChecksEachFieldAndEachNullAndLengthBoundary()
        {
            var rows = NativePaletteFixture.Rows();
            Assert.IsFalse(VbeNativePaletteState.Equal(null, rows));
            Assert.IsFalse(VbeNativePaletteState.Equal(rows, null));
            Assert.IsFalse(VbeNativePaletteState.Equal(rows, new Row[0]));
            Assert.IsTrue(VbeNativePaletteState.Equal(new Row[0], new Row[0]));
            foreach (var change in new Action<Row>[] { r => r.Name = "Other", r => r.Foreground++, r => r.Background--, r => r.Indicator++ })
            {
                var other = NativePaletteFixture.Rows(); change(other[0]);
                Assert.IsFalse(VbeNativePaletteState.Equal(rows, other));
            }
            var absent = NativePaletteFixture.Rows(); absent[0] = null;
            Assert.IsFalse(VbeNativePaletteState.Equal(absent, rows));
            Assert.IsFalse(VbeNativePaletteState.Equal(rows, absent));
            Assert.IsTrue(VbeNativePaletteState.Equal(rows, NativePaletteFixture.Rows()));
        }

        [TestMethod]
        public void DarkMappingIsExactAndDoesNotMutateOriginal()
        {
            var rows = NativePaletteFixture.Rows(); var dark = VbeNativePaletteState.Dark(rows);
            CollectionAssert.AreEqual(new[] { 8, 16, 13, 1, 16, 3, 4, 8, 8, 1 }, dark.Select(r => r.Foreground).ToArray());
            CollectionAssert.AreEqual(new[] { 1, 2, 1, 15, 5, 1, 1, 1, 1, 15 }, dark.Select(r => r.Background).ToArray());
            CollectionAssert.AreEqual(rows.Select(r => r.Name).ToArray(), dark.Select(r => r.Name).ToArray());
            CollectionAssert.AreEqual(rows.Select(r => r.Indicator).ToArray(), dark.Select(r => r.Indicator).ToArray());
            Assert.IsTrue(VbeNativePaletteState.Equal(rows, NativePaletteFixture.Rows()));
        }

        [TestMethod]
        public void RecoveryLoadRejectsMissingOversizedNullAndMalformedFiles()
        {
            using (var fixture = new NativePaletteFixture())
            {
                Assert.IsNull(VbeNativePaletteState.Load(fixture.PathName, "7.1"));
                File.WriteAllText(fixture.PathName, new string(' ', 32769));
                Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.Load(fixture.PathName, "7.1"));
                File.WriteAllText(fixture.PathName, "null");
                Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.Load(fixture.PathName, "7.1"));
                File.WriteAllText(fixture.PathName, "{");
                Assert.ThrowsException<ArgumentException>(() => VbeNativePaletteState.Load(fixture.PathName, "7.1"));
            }
        }

        [TestMethod]
        public void SavedStateValidationRejectsSchemaVersionAndBothInvalidPalettes()
        {
            foreach (var corrupt in new Action<VbeNativePaletteState>[] { s => s.Schema = 2, s => s.VbeVersion = "6.0",
                s => s.Original = null, s => s.Applied = null, s => s.Applied[0].Foreground = 2 })
            {
                var state = NativePaletteFixture.State(); corrupt(state);
                Assert.ThrowsException<InvalidDataException>(() => state.Validate("7.1"));
                using (var fixture = new NativePaletteFixture())
                {
                    if (state.VbeVersion != "6.0")
                    { Assert.ThrowsException<InvalidDataException>(() => state.SaveNew(fixture.PathName)); Assert.IsFalse(File.Exists(fixture.PathName)); }
                }
            }
        }
    }
}
