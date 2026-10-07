using System;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Unit
{
    internal sealed class NativePaletteFixture : IDisposable
    {
        internal readonly string DirectoryPath = Path.Combine(Path.GetTempPath(), "VBAi-palette-matrix-" + Guid.NewGuid().ToString("N"));
        internal string PathName => Path.Combine(DirectoryPath, "palette.json");
        internal VbeNativePaletteState.ColorRow[] Current = Rows();
        internal int Visits, Updates;
        internal bool RejectCommit;
        internal Exception Failure;
        internal NativePaletteFixture() { Directory.CreateDirectory(DirectoryPath); }
        internal static VbeNativePaletteState.ColorRow[] Rows() => Enumerable.Range(0, 10).Select(i => new VbeNativePaletteState.ColorRow
        { Name = "Category " + i, Foreground = i, Background = 16 - i, Indicator = i + 1 }).ToArray();
        internal static VbeNativePaletteState State()
        {
            var rows = Rows();
            return new VbeNativePaletteState { VbeVersion = "7.1", Original = rows, Applied = VbeNativePaletteState.Dark(rows) };
        }
        internal VbeNativePaletteState.ColorRow[] Visit(Func<VbeNativePaletteState.ColorRow[], VbeNativePaletteState.ColorRow[]> update)
        {
            Visits++;
            if (Failure != null) throw Failure;
            var desired = update(Current);
            if (desired != null) { Updates++; if (!RejectCommit) Current = desired; }
            return Current;
        }
        internal void Change(bool enabled) => VbeNativePalette.Change("7.1", enabled, PathName, Visit);
        public void Dispose() { Directory.Delete(DirectoryPath, true); }
    }
}
