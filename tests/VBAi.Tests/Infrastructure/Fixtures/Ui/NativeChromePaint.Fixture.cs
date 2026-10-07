using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    internal sealed class NativeChromePaintFixture : IDisposable
    {
        private readonly Dictionary<FieldInfo, object> fields = new Dictionary<FieldInfo, object>();
        private readonly Action<string, string> log = LoadLog.AppendText;
        internal readonly List<string> Messages = new List<string>();
        internal readonly IntPtr Window, Child = new IntPtr(123);
        internal IntPtr Dc;
        internal int Releases, ClientAcquires, WindowAcquires, Allocations;
        internal bool BoundsAvailable = true, ClientAvailable = true, OriginAvailable = true, ComboAvailable = true,
            CursorAvailable = true, CursorConversion = true, Enabled = true, HasChild, ChildAvailable = true;
        internal int Width = 24, Height = 16, LeftEdge = 2, RightEdge = 2, BottomEdge = 2, Caption = 3;
        internal int ButtonLeft = 12, ButtonTop, ButtonRight = 24, ButtonBottom = 16;
        internal uint ButtonState;
        internal int PointerX = 20, PointerY = 8;
        internal bool MissingDc, Reenter;
        internal Exception AllocationFailure;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateWindowEx(int exStyle, string className,
            string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);

        internal NativeChromePaintFixture()
        {
            foreach (var field in typeof(VbeNativeChrome).GetFields(BindingFlags.Static | BindingFlags.NonPublic))
                if (!field.IsInitOnly && !field.IsLiteral) fields[field] = field.GetValue(null);
            Window = CreateWindowEx(0, "STATIC", "Synthetic Chrome canvas", 0, 0, 0, 24, 16, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, Window);
            VbeNativeChrome.WindowBounds = Bounds; VbeNativeChrome.ClientBounds = Client;
            VbeNativeChrome.ToScreen = Origin; VbeNativeChrome.ToClient = Convert;
            VbeNativeChrome.PointerPosition = Pointer; VbeNativeChrome.ComboInformation = Combo;
            VbeNativeChrome.WindowEnabled = window => { Assert.AreEqual(Window, window); return Enabled; };
            VbeNativeChrome.RelatedWindow = (window, relation) => { Assert.AreEqual(Window, window); Assert.AreEqual((uint)5, relation); return HasChild ? Child : IntPtr.Zero; };
            VbeNativeChrome.AcquireWindowDc = window => { Assert.AreEqual(Window, window); WindowAcquires++; return MissingDc ? IntPtr.Zero : Dc; };
            VbeNativeChrome.AcquireClientDc = window => { Assert.AreEqual(Window, window); ClientAcquires++; return MissingDc ? IntPtr.Zero : Dc; };
            VbeNativeChrome.ReleaseWindowDc = (window, dc) => { Assert.AreEqual(Window, window); Assert.AreEqual(Dc, dc); Releases++; return 1; };
            VbeNativeChrome.CreateChromeBitmap = (width, height) =>
            {
                Allocations++; if (Reenter) VbeNativeChrome.Paint(Window, true, Dc);
                if (AllocationFailure != null) throw AllocationFailure;
                return new Bitmap(width, height, PixelFormat.Format32bppRgb);
            };
            LoadLog.AppendText = (path, message) => Messages.Add(message);
        }
        private bool Bounds(IntPtr window, out VbeNativeChrome.Rect rectangle)
        {
            if (window == Child) { rectangle = new VbeNativeChrome.Rect { Top = 6 }; return ChildAvailable; }
            Assert.AreEqual(Window, window); rectangle = new VbeNativeChrome.Rect { Right = Width, Bottom = Height }; return BoundsAvailable;
        }
        private bool Client(IntPtr window, out VbeNativeChrome.Rect rectangle)
        { Assert.AreEqual(Window, window); rectangle = new VbeNativeChrome.Rect { Right = Width - LeftEdge - RightEdge, Bottom = Height - Caption - BottomEdge }; return ClientAvailable; }
        private bool Origin(IntPtr window, ref VbeNativeChrome.Point point)
        { Assert.AreEqual(Window, window); point.X = LeftEdge; point.Y = Caption; return OriginAvailable; }
        private bool Convert(IntPtr window, ref VbeNativeChrome.Point point) { Assert.AreEqual(Window, window); return CursorConversion; }
        private bool Pointer(out VbeNativeChrome.Point point) { point = new VbeNativeChrome.Point { X = PointerX, Y = PointerY }; return CursorAvailable; }
        private bool Combo(IntPtr window, ref VbeNativeChrome.ComboInfo information)
        {
            Assert.AreEqual(Window, window); Assert.IsTrue(information.Size > 0);
            information.Button = new VbeNativeChrome.Rect { Left = ButtonLeft, Top = ButtonTop, Right = ButtonRight, Bottom = ButtonBottom };
            information.ButtonState = ButtonState; return ComboAvailable;
        }
        internal Bitmap Canvas(Color initial, Action action) => NativeChromeCanvas.Paint(initial, dc => { Dc = dc; action(); });
        public void Dispose()
        {
            DestroyWindow(Window); foreach (var pair in fields) pair.Key.SetValue(null, pair.Value); LoadLog.AppendText = log;
        }
    }
}
