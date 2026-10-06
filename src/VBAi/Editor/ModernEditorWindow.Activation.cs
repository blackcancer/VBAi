using System;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns the modern editor window state and operations.</summary>
    internal sealed partial class ModernEditorWindow
    {

        /// <summary>Reads only the identity-matched native code pane currently activated in the owning VBE.</summary>
        internal Func<IEditorModule> ReadActiveModule;

        /// <summary>Maintains the last followed native module state for modern editor window.</summary>
        private IEditorModule lastFollowedNativeModule;

        /// <summary>Follows native activation independently of its mouse, accessibility or programmatic origin.</summary>
        /// <returns>task produced by the operation for follow native activation on modern editor window.</returns>
        internal Task FollowNativeActivation()
        {
            if (ReadActiveModule == null || !WorkspaceHosted || !Ready || busy || closing || IsDisposed ||
                VbeDebugInspection.IsActive || debugCommands.CurrentCount == 0) return Task.CompletedTask;
            return VbeUiTask.Run(FollowNativeActivationCore);
        }

        /// <summary>Handles follow native activation core for modern editor window.</summary>
        /// <returns>task&lt;bool&gt; produced by the operation for follow native activation core on modern editor window.</returns>
        private async Task<bool> FollowNativeActivationCore()
        {
            // Posting to the owning STA yields; another operation can acquire the
            // editor or close it before this callback starts.
            if (busy || closing || IsDisposed || !Ready || !WorkspaceHosted ||
                VbeDebugInspection.IsActive || debugCommands.CurrentCount == 0) return false;
            busy = true;
            try
            {
                var module = ReadActiveModule();
                if (module == null) { lastFollowedNativeModule = null; return false; }
                // Native activation follows transitions. A Monaco tab selection must remain
                // usable while the unchanged backing VBE pane is still active.
                if (!Visible || SameEditorModule(lastFollowedNativeModule, module)) return false;
                if (SameEditorModule(Current?.Module, module)) { lastFollowedNativeModule = module; return false; }
                await CaptureDocuments(); // Preserve the latest unsynchronized draft before switching tabs.
                if (closing || IsDisposed || !Ready || !Visible || debugCommands.CurrentCount == 0 || VbeDebugInspection.IsActive)
                    return false;
                // WebView capture yields. Never open a stale component after navigation,
                // a designer activation, protection or mode changes during that boundary.
                if (!SameEditorModule(module, ReadActiveModule())) return false;
                await OpenModuleCore(module, false);
                lastFollowedNativeModule = module;
                return true;
            }
            catch (Exception error)
            {
                LoadLog.Write("Monaco active module observation refused: " + error.GetType().Name);
                return false;
            }
            finally { busy = false; }
        }

        /// <summary>Compares exact adapter/component identities; mutable display names never identify a document.</summary>
        /// <param name="first">i editor module that supplies the first for this operation.</param>
        /// <param name="second">i editor module that supplies the second for this operation.</param>
        /// <returns>Boolean indicating the result of the check for same editor module on modern editor window.</returns>
        private static bool SameEditorModule(IEditorModule first, IEditorModule second) =>
            first != null && second != null && (ReferenceEquals(first, second) ||
                (first is EditorVbeModule native && second is EditorVbeModule other && native.IsComponent(other.Component)));
    }
}
