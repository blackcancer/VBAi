using System;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;

namespace VBAi.Tests.Infrastructure
{
    internal sealed class TechnicalUiMetadataScope : IDisposable
    {
        private readonly Func<Assembly> aboutAssembly = AboutWindow.MetadataAssembly, crashAssembly = CrashReport.MetadataAssembly;
        private readonly Func<bool> aboutBits = AboutWindow.ProcessIs64Bit, crashBits = CrashReport.ProcessIs64Bit;
        private readonly Func<Version> aboutVersion = AboutWindow.RuntimeVersion, crashVersion = CrashReport.RuntimeVersion;
        private readonly Func<string> directory = CrashReport.ReportDirectory;
        private readonly Func<Exception, StackFrame[]> frames = CrashReport.FrameSnapshot;
        internal static Assembly WithoutInformation() => AssemblyBuilder.DefineDynamicAssembly(new AssemblyName("OwnedTechnicalMetadata" + Guid.NewGuid().ToString("N")) { Version = new Version(5, 6, 7, 8) }, AssemblyBuilderAccess.Run);
        public void Dispose() { AboutWindow.MetadataAssembly = aboutAssembly; AboutWindow.ProcessIs64Bit = aboutBits; CrashReport.MetadataAssembly = crashAssembly; CrashReport.ProcessIs64Bit = crashBits; CrashReport.FrameSnapshot = frames; AboutWindow.RuntimeVersion = aboutVersion; CrashReport.RuntimeVersion = crashVersion; CrashReport.ReportDirectory = directory; }
    }
    internal sealed class OwnedTechnicalStackFrame : StackFrame
    {
        private readonly MethodBase method;
        internal OwnedTechnicalStackFrame(MethodBase method) { this.method = method; }
        public override MethodBase GetMethod() => method;
    }
}