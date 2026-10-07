using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Retains source font interfaces across one import, without retaining a removed designer.</summary>
        internal GitSourceFontLease CaptureGitSourceFonts(string form, string layout, IDictionary<string, object> report)
        {
            var lease = new GitSourceFontLease();
            try
            {
                WithGitLayoutDesigner(form, (component, designer) =>
                {
                    lease.Form = ((dynamic)designer).Font;
                    report["SourceFormFontIdentity"] = DescribeGitFontPersistence(lease.Form);
                    if (layout != "FrameMultiPage") return;
                    object controls = null, frame = null;
                    try
                    {
                        controls = ((dynamic)designer).Controls;
                        frame = ((dynamic)controls).Item("QualificationExtra");
                        lease.Frame = ((dynamic)frame).Font;
                        report["SourceFrameFontIdentity"] = DescribeGitFontPersistence(lease.Frame);
                    }
                    finally { Release(frame); Release(controls); }
                });
                return lease;
            }
            catch { lease.Dispose(); throw; }
        }

        /// <summary>Assigns each separately retained source font once to the current owned imported designer.</summary>
        internal void RestoreGitSourceFonts(string form, string layout, GitSourceFontLease lease)
        {
            if (lease == null || lease.Form == null || (layout == "FrameMultiPage" && lease.Frame == null))
                throw new InvalidOperationException("An intact source-font lease is required before assignment.");
            WithGitLayoutDesigner(form, (component, designer) =>
            {
                ((dynamic)designer).Font = lease.Form;
                if (layout != "FrameMultiPage") return;
                object controls = null, frame = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    frame = ((dynamic)controls).Item("QualificationExtra");
                    ((dynamic)frame).Font = lease.Frame;
                }
                finally { Release(frame); Release(controls); }
            });
        }

        /// <summary>Reads the persistence class and exact descriptor without invoking font metric getters.</summary>
        private static object DescribeGitFontPersistence(object font)
        {
            IStream stream = null;
            try
            {
                var persisted = (GitDiagnosticPersistStream)font;
                Guid id; persisted.GetClassID(out id);
                Marshal.ThrowExceptionForHR(CreateGitFontStream(IntPtr.Zero, true, out stream));
                persisted.Save(stream, false);
                System.Runtime.InteropServices.ComTypes.STATSTG state; stream.Stat(out state, 1);
                if (state.cbSize < 0 || state.cbSize > 4096)
                    throw new InvalidOperationException("Source font descriptor exceeds the diagnostic bound.");
                byte[] bytes = new byte[(int)state.cbSize];
                stream.Seek(0, 0, IntPtr.Zero); stream.Read(bytes, bytes.Length, IntPtr.Zero);
                return new { ClassId = id.ToString("D"), DescriptorHex = BitConverter.ToString(bytes), MetricGetters = 0 };
            }
            finally { Release(stream); }
        }

        /// <summary>Balances acquired font references once on the fixture STA, including failure paths.</summary>
        internal sealed class GitSourceFontLease : IDisposable
        {
            internal object Form, Frame;
            public void Dispose()
            {
                object frame = Frame, form = Form; Frame = Form = null;
                try { Release(frame); } finally { Release(form); }
            }
        }
    }
}
