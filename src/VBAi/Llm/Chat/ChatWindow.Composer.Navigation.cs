using System;
using System.ComponentModel;
using System.Windows;

namespace VBAi
{
    /// <summary>Opens referenced code without replacing the conversation draft.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>The project whose fresh modules are displayed for navigation, rather than insertion.</summary>
        private VbeChatReference referenceNavigationProject;

        /// <summary>Observes the popup's synchronous open state instead of its deferred Closed event.</summary>
        private DependencyPropertyDescriptor referencePopupOpenDescriptor;

        /// <summary>Clears navigation when the popup closes; a delayed close from an earlier opening cannot clear a newer picker.</summary>
        /// <param name="sender">The popup whose dependency property changed.</param>
        /// <param name="args">Dependency-property change notification.</param>
        private void ReferencePopupOpenChanged(object sender, EventArgs args)
        {
            if (!referencePopup.IsOpen) referenceNavigationProject = null;
        }

        /// <summary>Browses a project's modules or navigates to code, surfacing bridge refusals.</summary>
        /// <param name="reference">The selected project, module or declaration.</param>
        private void NavigateReference(VbeChatReference reference)
        {
            try
            {
                if (reference == null) throw new ArgumentNullException(nameof(reference));
                if (reference.Module == null) { BrowseReferenceProject(reference); return; }
                var response = referenceIndex.Navigate(reference);
                if (!response.Ok) SetStatus(UiText.Get("Unable to navigate: ") + response.Error);
            }
            catch (Exception error) { SetStatus(UiText.Get("Unable to navigate: ") + error.Message); }
        }

        /// <summary>Shows current modules in the existing suggestions view, preserving text, selection and context chips.</summary>
        /// <param name="reference">A project reference admitted by the live VBE resolver.</param>
        private void BrowseReferenceProject(VbeChatReference reference)
        {
            HideReferences();
            var modules = referenceIndex.ModulesForProject(reference);
            SetSuggestionTargets(modules);
            referenceList.SelectedIndex = modules.Length > 0 ? 0 : -1;
            referenceList.Visible = modules.Length > 0;
            referenceView.Width = prompt.ActualWidth > 0 ? (int)Math.Max(240, Math.Min(480, prompt.ActualWidth - 14)) : 350;
            referenceList.Height = Math.Max(30, Math.Min(240, modules.Length * referenceList.ItemHeight));
            ((FrameworkElement)referencePopup.Child).Width = referenceView.Width;
            referenceStatus.MaximumSize = new System.Drawing.Size(Math.Max(200, referenceView.Width - 24), 0);
            referenceStatus.Text = modules.Length == 0 ? UiText.Get("No matching target") :
                UiText.Get("Select a module or procedure in this project to open its code.");
            referencePopup.HorizontalOffset = 0;
            referencePopup.VerticalOffset = prompt.GetRectFromCharacterIndex(prompt.CaretIndex).Bottom + 3;
            referenceNavigationProject = reference;
            referencePopup.IsOpen = true;
        }
    }
}
