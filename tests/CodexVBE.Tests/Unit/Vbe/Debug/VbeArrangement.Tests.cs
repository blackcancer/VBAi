using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeArrangementTests
    {
        private static EditorWindowBounds Rect(int left, int top, int width = 300, int height = 200)
        {
            return new EditorWindowBounds { Left = left, Top = top, Width = width, Height = height };
        }

        [TestMethod]
        public void TilingRequiresTheRequestedOrientation()
        {
            var columns = new[] { Rect(0, 0), Rect(300, 0), Rect(600, 0) };
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_vertical", columns));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_horizontal", columns));
            var lines = new[] { Rect(0, 0), Rect(0, 200), Rect(0, 400) };
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_horizontal", lines));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", lines));
        }

        [TestMethod]
        public void TilingRejectsOverlapsGapsAndUnequalSizes()
        {
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(280, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(500, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(300, 0, 150) }));
            Assert.IsTrue(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0), Rect(301, 0, 301) }));
        }

        [TestMethod]
        public void CascadeRequiresOverlappingDiagonalStepsRegardlessOfIdentityOrder()
        {
            Assert.IsTrue(VbeDebug.VerifyArrangement("cascade", new[] { Rect(52, 52), Rect(0, 0), Rect(26, 26) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(26, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(300, 200) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0), Rect(0, 0) }));
        }

        [TestMethod]
        public void InvalidOrMaximizedWindowsCannotVerifyAnArrangement()
        {
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", null));
            Assert.IsFalse(VbeDebug.VerifyArrangement("cascade", new[] { Rect(0, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { Rect(0, 0, 0), Rect(300, 0) }));
            var maximized = Rect(0, 0); maximized.State = 2;
            Assert.IsFalse(VbeDebug.VerifyArrangement("tile_vertical", new[] { maximized, Rect(300, 0) }));
            Assert.IsFalse(VbeDebug.VerifyArrangement("unknown", new[] { Rect(0, 0), Rect(300, 0) }));
        }
    }
}
