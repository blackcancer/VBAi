using System;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Follows identity-verified VBE module activation into the matching Monaco document.</summary>
    internal sealed partial class ModernEditorWindow
    {

        /// <summary>Reads only the identity-matched native code pane currently activated in the owning VBE.</summary>
        internal Func<IEditorModule> ReadActiveModule;

        /// <summary>Last native module followed, used to distinguish native activation transitions from Monaco tab changes.</summary>
        private IEditorModule lastFollowedNativeModule;

        /// <summary>Follows native activation independently of its mouse, accessibility or programmatic origin.</summary>
        /// <returns>A task that completes after the native activation has been observed and, when needed, its document opened.</returns>
        internal Task FollowNativeActivation()
        {
            if (ReadActiveModule == null || !WorkspaceHosted || !Ready || busy || closing || IsDisposed ||
                VbeDebugInspection.IsActive || debugCommands.CurrentCount == 0) return Task.CompletedTask;
            return VbeUiTask.Run(FollowNativeActivationCore);
        }

        /// <summary>Rechecks editor state after dispatch to the owning STA, captures pending drafts, then follows the still-active module.</summary>
        /// <returns><see langword="true"/> when a different active module was opened; otherwise <see langword="false"/>.</returns>
        private async Task<bool> FollowNativeActivationCore()
        {
            // Posting to the owning STA yields; another operation can acquire the
            // editor or close it before this callback starts.
            if (busy || closing || IsDisposed || !Ready || !WorkspaceHosted ||
                VbeDebugInspection.IsActive || debugCommands.CurrentCount == 0) return false;
            busy = true;
            IEditorModule module = null;
            try
            {
                module = ReadActiveModule();
                if (module == null) { lastFollowedNativeModule = null; return false; }
                // Native activation follows transitions. A Monaco tab selection must remain
                // usable while the unchanged backing VBE pane is still active.
                if (!Visible || SameEditorModule(lastFollowedNativeModule, module)) return false;
                if (SameEditorModule(Current?.Module, module)) { lastFollowedNativeModule = Current.Module; return false; }
                await CaptureDocuments(); // Preserve the latest unsynchronized draft before switching tabs.
                if (closing || IsDisposed || !Ready || !Visible || debugCommands.CurrentCount == 0 || VbeDebugInspection.IsActive)
                    return false;
                // WebView capture yields. Never open a stale component after navigation,
                // a designer activation, protection or mode changes during that boundary.
                var verified = ReadActiveModule();
                try { if (!SameEditorModule(module, verified)) return false; }
                finally { if (!ReferenceEquals(verified, module)) ReleaseUnadoptedModule(verified); }
                var document = await OpenModuleCore(module, false);
                lastFollowedNativeModule = document.Module;
                return true;
            }
            catch (Exception error)
            {
                LoadLog.Write("Monaco active module observation refused: " + error.GetType().Name);
                return false;
            }
            finally { try { ReleaseUnadoptedModule(module); } finally { busy = false; } }
        }

        /// <summary>Reports whether a live document has adopted this exact adapter instance.</summary>
        /// <param name="module">Possibly null adapter whose ownership is being checked.</param>
        /// <returns>True only for an exact adopted adapter, not for a matching component probe.</returns>
        internal bool OwnsModule(IEditorModule module) =>
            module != null && System.Linq.Enumerable.Any(documents.Values, document => ReferenceEquals(document.Module, module));

        /// <summary>Releases a transient adapter only when no open document has adopted it.</summary>
        /// <param name="module">Possibly null incoming or validation-only adapter.</param>
        internal void ReleaseUnadoptedModule(IEditorModule module)
        { if (!OwnsModule(module)) (module as EditorVbeModule)?.Dispose(); }

        /// <summary>Compares exact adapter/component identities; mutable display names never identify a document.</summary>
        /// <param name="first">First adapter module identity.</param>
        /// <param name="second">Second adapter module identity.</param>
        /// <returns><see langword="true"/> when both references are the same adapter object or wrap the same VBE component.</returns>
        private static bool SameEditorModule(IEditorModule first, IEditorModule second) =>
            first != null && second != null && (ReferenceEquals(first, second) ||
                (first is EditorVbeModule native && second is EditorVbeModule other && native.IsComponent(other.Component)));
    }
}
