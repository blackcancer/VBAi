namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        public sealed class ExcelDocumentObject { public object Application { get; set; } }
        public sealed class ExcelApplicationObject { public long Hwnd { get; set; } }

        /// <summary>Simulates independently acquired NativeOM window and application interfaces.</summary>
        public sealed class ExcelResolutionObject
        {
            public System.Func<object> ReadApplication;
            public System.Func<object> ReadHandle;
            public object Application => ReadApplication();
            public object Hwnd => ReadHandle();
        }
    }
}
