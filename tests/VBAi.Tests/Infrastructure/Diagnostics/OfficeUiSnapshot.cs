using System;
using System.Windows.Automation;

namespace VBAi.Tests.Integration
{
    /// <summary>Fetches a fresh set of control properties in one UI Automation request.</summary>
    internal static class OfficeUiSnapshot
    {
        /// <summary>Retains action patterns but caches only the inspected element, never a changing subtree.</summary>
        internal static CacheRequest CreateRequest()
        {
            var request = new CacheRequest
            {
                TreeScope = TreeScope.Element,
                TreeFilter = Automation.RawViewCondition,
                AutomationElementMode = AutomationElementMode.Full
            };
            request.Add(AutomationElement.AutomationIdProperty);
            request.Add(AutomationElement.NameProperty);
            request.Add(AutomationElement.ProcessIdProperty);
            request.Add(AutomationElement.NativeWindowHandleProperty);
            request.Add(AutomationElement.ControlTypeProperty);
            request.Add(AutomationElement.IsEnabledProperty);
            request.Add(AutomationElement.IsOffscreenProperty);
            return request;
        }

        /// <summary>Checks ownership on both sides of a potentially blocking read; no action or retry occurs here.</summary>
        internal static T ReadOnce<T>(Action requireOwner, Func<T> read)
        {
            if (read == null) throw new ArgumentNullException(nameof(read));
            requireOwner?.Invoke();
            T result = read();
            requireOwner?.Invoke();
            return result;
        }

        /// <summary>Returns the updated peer; the input peer's earlier cached values are never reused.</summary>
        internal static AutomationElement Refresh(AutomationElement element, Action requireOwner = null)
        {
            if (element == null) throw new ArgumentNullException(nameof(element));
            return ReadOnce(requireOwner, () => element.GetUpdatedCache(CreateRequest()));
        }

        /// <summary>Inspects a native HWND peer's identity without traversing the host's accessibility subtree.</summary>
        internal static AutomationElement RefreshIdentity(AutomationElement element, Action requireOwner = null)
        {
            if (element == null) throw new ArgumentNullException(nameof(element));
            var request = new CacheRequest
            {
                TreeScope = TreeScope.Element,
                TreeFilter = Automation.RawViewCondition,
                AutomationElementMode = AutomationElementMode.Full
            };
            request.Add(AutomationElement.AutomationIdProperty);
            request.Add(AutomationElement.ProcessIdProperty);
            request.Add(AutomationElement.NativeWindowHandleProperty);
            return ReadOnce(requireOwner, () => element.GetUpdatedCache(request));
        }

        /// <summary>Returns peers with the same property snapshot as a direct refresh.</summary>
        internal static AutomationElementCollection FindAll(AutomationElement root, TreeScope scope,
            Condition condition, Action requireOwner = null)
        {
            if (root == null) throw new ArgumentNullException(nameof(root));
            return ReadOnce(requireOwner, () =>
            {
                using (CreateRequest().Activate()) return root.FindAll(scope, condition);
            });
        }
    }
}
