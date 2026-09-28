namespace CodexVBE.Tests.Infrastructure
{
    using System;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed class VbeToolMode { public int Mode { get; set; } }
    public sealed class VbeToolSignature { public string CertificateName { get; set; } public bool UnsignedVerified { get; set; } }
    public sealed class VbeToolPersistence { public bool Saved { get; set; } }

    internal static class VbeToolBoundaryFixture
    {
        internal static void Configure(VbeToolNativeBoundary native)
        {
            native.Capture = stack => new { Native = "capture", Stack = stack };
            native.ReadDebugDialog = () => new { Native = "dialog" };
            native.ChangeDebugItem = request => new { Native = "item", request.Action };
            native.RespondDebugDialog = request => new { Native = "respond", request.Button };
            native.ExecuteImmediate = text => new { Native = "immediate", Text = text };
            native.EnsureNoCompileDialog = () => { };
            native.AwaitCompileDialog = completed => { Assert.IsTrue(completed.Wait(5000)); return null; };
            native.CompleteAddWatch = request => new { Native = "add_watch" };
            native.SelectWatch = request => new { Native = "select_watch" };
            native.CompleteEditWatch = request => new { Native = "edit_watch" };
            native.CompleteQuickWatch = request => new { Native = "quick_watch" };
            native.EnsureNoDebugOptionsDialog = () => { };
            native.ReadVbeOptions = () => new { Native = "vbe_options" };
            native.ReadDebugOptions = () => new { Native = "debug_options" };
            native.EnsureNoSignatureDialog = () => { };
            native.ReadSignatureDialog = project => new { Native = "signature_dialog", Project = project };
            native.CompleteProjectSignature = (project, thumbprint, name, unsigned) =>
                new { Native = "signature", Project = project, Thumbprint = thumbprint, CertificateName = name, UnsignedVerified = unsigned };
            native.VerifyWatchRemoved = request => new { Native = "remove_watch" };
        }

        internal static Response Execute(Request request)
        {
            if (request == null) return Response.Failure("request is null");
            if (request.Command == "debug_state") return Response.Success(new VbeToolMode { Mode = 2 });
            if (request.Command == "sign_project") return Response.Success(new VbeToolSignature { CertificateName = "Disposable", UnsignedVerified = true });
            return Response.Success(new { Command = request.Command });
        }
    }
}
