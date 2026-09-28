using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeEditorWindows
    {
        public object WindowLayout(string caption, int type)
        {
            dynamic window = FindExactWindow(caption, type);
            dynamic snapshot = WindowSnapshot(window, null);
            dynamic linkage = WindowLinkage(caption, type);
            var members = new List<object>();
            if (type == 11 || type == 12)
            {
                try { foreach (dynamic member in window.LinkedWindows) members.Add(new { Caption = (string)member.Caption, Type = (int)member.Type }); }
                catch (Exception error) { snapshot.Errors["LinkedWindows"] = error.Message; }
            }
            string version = null;
            if (snapshot.Errors.Count == 0 && linkage.Errors.Count == 0)
            {
                string state = new JavaScriptSerializer().Serialize(new { Window = snapshot.Properties, Linkage = linkage.Properties, FrameMembers = members });
                using (var sha = SHA256.Create()) version = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(state))).Replace("-", "").ToLowerInvariant();
            }
            return new { Window = snapshot, Linkage = linkage, FrameMembers = members, WindowVersion = version };
        }

        public object SetWindowState(Request request)
        {
            int desired;
            switch (request.Action)
            {
                case "restore": desired = 0; break;
                case "minimize": desired = 1; break;
                case "maximize": desired = 2; break;
                default: throw new ArgumentException("Use restore, minimize or maximize.");
            }
            dynamic target = FindExactWindow(request.WindowCaption, request.WindowType);
            object before = CheckedLayout(request.WindowCaption, request.WindowType, request.ExpectedWindowVersion);
            if (target.LinkedWindowFrame != null)
                throw new InvalidOperationException("Change the state of the containing frame, not a linked pane.");
            int original = (int)target.WindowState;
            if (original == desired) return new { Applied = false, Verified = true, Before = before, After = before, PersistenceVerified = false };
            string error = null;
            try { target.WindowState = desired; } catch (Exception ex) { error = ex.Message; }
            object after = null; string readError = null; int? actual = null;
            try
            {
                actual = (int)target.WindowState;
                // Native code-window captions change when leaving maximized state.
                after = WindowLayout((string)target.Caption, (int)target.Type);
            }
            catch (Exception ex) { readError = ex.Message; }
            bool verified = error == null && readError == null && actual == desired;
            return new { Applied = error == null ? (bool?)true : null, Verified = verified, VerificationPending = !verified,
                Before = before, After = after, ActualState = actual, NativeError = error, ReadbackError = readError,
                PersistenceVerified = false, NextRead = verified ? null : "vbe_windows" };
        }

        public object SetWindowBounds(Request request)
        {
            var values = new[] { request.Left, request.Top, request.Width, request.Height };
            if (values.Any(x => double.IsNaN(x) || double.IsInfinity(x) || x != Math.Truncate(x) || x < -32768 || x > 32767) || request.Width < 80 || request.Height < 60)
                throw new ArgumentException("Bounds require integer native coordinates, with width >= 80 and height >= 60.");
            dynamic target = FindExactWindow(request.WindowCaption, request.WindowType);
            dynamic before = CheckedLayout(request.WindowCaption, request.WindowType, request.ExpectedWindowVersion);
            if ((int)target.WindowState != 0) throw new InvalidOperationException("Restore the window to normal state before changing its bounds.");
            dynamic frame = target.LinkedWindowFrame;
            if (frame != null) throw new InvalidOperationException("Bounds target a standalone window; detach the pane and target its frame when applicable.");
            int left = (int)target.Left, top = (int)target.Top, width = (int)target.Width, height = (int)target.Height;
            try
            {
                target.Left = (int)request.Left; target.Top = (int)request.Top;
                target.Width = (int)request.Width; target.Height = (int)request.Height;
                if ((int)target.Left != (int)request.Left || (int)target.Top != (int)request.Top || (int)target.Width != (int)request.Width || (int)target.Height != (int)request.Height)
                    throw new InvalidOperationException("The host constrained or rejected the requested bounds.");
            }
            catch (Exception error)
            {
                try
                {
                    target.Left = left; target.Top = top; target.Width = width; target.Height = height;
                    if ((int)target.Left != left || (int)target.Top != top || (int)target.Width != width || (int)target.Height != height)
                        throw new InvalidOperationException("Original bounds could not be verified.");
                }
                catch (Exception rollback) { throw new InvalidOperationException(error.Message + " Bounds rollback failed: " + rollback.Message, error); }
                throw new InvalidOperationException(error.Message + " Original bounds restored.", error);
            }
            return new { Applied = true, Verified = true, Before = (object)before,
                After = WindowLayout(request.WindowCaption, request.WindowType), PersistenceVerified = false };
        }

        public object LinkWindow(Request request)
        {
            if (request.Action != "link" && request.Action != "unlink") throw new ArgumentException("Use link or unlink.");
            if (!(request.WindowType >= 2 && request.WindowType <= 7) && request.WindowType != 10 && request.WindowType != 15)
                throw new ArgumentException("Linking targets native tool panes, not code windows, designers or frame containers.");
            dynamic target = FindExactWindow(request.WindowCaption, request.WindowType);
            object before = CheckedLayout(request.WindowCaption, request.WindowType, request.ExpectedWindowVersion);
            dynamic oldFrame = target.LinkedWindowFrame;
            dynamic destination = null;
            if (request.Action == "link")
            {
                if (request.WindowCaption == request.TargetWindowCaption && request.WindowType == request.TargetWindowType)
                    throw new ArgumentException("A window cannot be linked to itself.");
                dynamic anchor = FindExactWindow(request.TargetWindowCaption, request.TargetWindowType);
                CheckedLayout(request.TargetWindowCaption, request.TargetWindowType, request.ExpectedTargetWindowVersion);
                destination = request.TargetWindowType == 11 || request.TargetWindowType == 12 ? anchor : anchor.LinkedWindowFrame;
                if (destination == null) throw new InvalidOperationException("The target has no linked frame. Select an existing frame or the VBE main window.");
                // Reading the collection before mutation detects unsupported host surfaces.
                int count = 0; foreach (dynamic member in destination.LinkedWindows) count++;
            }
            if (request.Action == "unlink" && oldFrame == null)
                return new { Applied = false, Verified = true, Before = before, After = before, PersistenceVerified = false };
            try
            {
                if (request.Action == "link") destination.LinkedWindows.Add(target);
                else oldFrame.LinkedWindows.Remove(target);
                dynamic currentFrame = target.LinkedWindowFrame;
                bool verified = request.Action == "link"
                    ? currentFrame != null && object.Equals((object)currentFrame, (object)destination) && FrameContains(destination, request.WindowCaption, request.WindowType)
                    : currentFrame == null || (!object.Equals((object)currentFrame, (object)oldFrame) && SingleMemberFrame(currentFrame, request.WindowCaption, request.WindowType));
                return new { Applied = (bool?)true, Verified = verified, VerificationPending = !verified,
                    Before = before, After = WindowLayout(request.WindowCaption, request.WindowType), PersistenceVerified = false };
            }
            catch (Exception error)
            {
                // VBIDE may destroy an emptied frame. Do not assume a stale frame RCW can
                // reconstruct the original topology after a partially completed native call.
                object after = null; string readError = null;
                try { after = WindowLayout(request.WindowCaption, request.WindowType); } catch (Exception read) { readError = read.Message; }
                return new { Applied = (bool?)null, Verified = false, VerificationPending = true,
                    Before = before, After = after, NativeError = error.Message, ReadbackError = readError,
                    PersistenceVerified = false, NextRead = "window_layout" };
            }
        }

        private object CheckedLayout(string caption, int type, string expected)
        {
            dynamic state = WindowLayout(caption, type);
            if (string.IsNullOrWhiteSpace(expected) || state.WindowVersion == null || !string.Equals(expected, (string)state.WindowVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The window layout changed or could not be read completely. Read window_layout again.");
            return state;
        }
        private static bool FrameContains(dynamic frame, string caption, int type)
        {
            int matches = 0;
            foreach (dynamic window in frame.LinkedWindows) if ((string)window.Caption == caption && (int)window.Type == type) matches++;
            return matches == 1;
        }
        private static bool SingleMemberFrame(dynamic frame, string caption, int type)
        {
            int count = 0; foreach (dynamic window in frame.LinkedWindows) count++;
            return count == 1 && FrameContains(frame, caption, type);
        }
    }
}
