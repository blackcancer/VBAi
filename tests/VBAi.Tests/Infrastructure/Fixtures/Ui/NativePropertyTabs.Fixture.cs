using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    internal sealed class NativePropertyTabsFixture : IDisposable
    {
        private readonly Dictionary<FieldInfo, object> saved = new Dictionary<FieldInfo, object>();
        internal readonly NativeThemeFixture Owner = new NativeThemeFixture();
        internal readonly VbeNativePropertyTabs Renderer;
        internal readonly IntPtr Window;
        internal readonly uint Thread;
        internal readonly List<uint> Tracking = new List<uint>();
        internal int Invalidations, Updates, Ends;
        internal bool TrackResult = true;
        internal NativePropertyTabsFixture()
        {
            foreach (var field in typeof(VbeNativePropertyTabs).GetFields(BindingFlags.Static | BindingFlags.NonPublic))
                if (!field.IsLiteral && !field.IsInitOnly) saved.Add(field, field.GetValue(null));
            Renderer = Owner.AddPropertyTabs();
            Window = (IntPtr)typeof(VbeNativePropertyTabs).GetField("window", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Renderer);
            Thread = VbeNativePropertyTabs.CurrentThread();
            VbeNativePropertyTabs.Invalidate = (window, rectangle, erase) => { Assert.AreEqual(Window, window); Assert.IsFalse(erase); Invalidations++; return true; };
            VbeNativePropertyTabs.Update = window => { Assert.AreEqual(Window, window); Updates++; return true; };
            VbeNativePropertyTabs.Track = (ref VbeNativePropertyTabs.TrackMouse state) => { Assert.AreEqual(Window, state.Window); Tracking.Add(state.Flags); return TrackResult; };
        }
        internal void ResetCallbacks()
        {
            foreach (var item in saved) item.Key.SetValue(null, item.Value);
        }
        internal static VbeNativePropertyTabs.Rect Bounds(int left = 0, int top = 0, int right = 240, int bottom = 60) =>
            new VbeNativePropertyTabs.Rect { Left = left, Top = top, Right = right, Bottom = bottom };
        internal bool Handle(uint message, IntPtr dc = default(IntPtr), long flags = 4) => Renderer.TryHandleMessage(message, dc, new IntPtr(flags), out _);
        internal void Hover(int x, int y) => Renderer.AfterNativeMessage(0x200, IntPtr.Zero, new IntPtr((y & 0xffff) << 16 | (x & 0xffff)));
        internal void Property(string name, object value) => typeof(VbeNativePropertyTabs).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Renderer, value);
        public void Dispose()
        {
            Renderer.Dispose(); ResetCallbacks(); Owner.Dispose();
        }
    }
}
