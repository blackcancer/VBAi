// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Blackcancer.ChatUI.Presentation
{
    public enum ChatReferenceNavigationState { Unavailable, Available, Stale }

    /// <summary>Detached captured text, not a live CAD/VBE object or a navigation authorization.</summary>
    public sealed class ChatReferenceSnapshot
    {
        public ChatReferenceSnapshot(string id, string label, string origin, IDictionary<string, string> metadata,
            string exactText, bool truncated, ChatReferenceNavigationState navigationState = ChatReferenceNavigationState.Unavailable,
            string status = null)
        {
            if (!Enum.IsDefined(typeof(ChatReferenceNavigationState), navigationState))
                throw new ArgumentOutOfRangeException(nameof(navigationState));
            Id = id ?? string.Empty; Label = label ?? string.Empty; Origin = origin ?? string.Empty;
            ExactText = exactText ?? string.Empty; Truncated = truncated;
            NavigationState = navigationState; Status = status ?? string.Empty;
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            if (metadata != null)
                foreach (var entry in metadata) values.Add(entry.Key, entry.Value ?? string.Empty);
            Metadata = new ReadOnlyDictionary<string, string>(values);
        }
        public string Id { get; }
        public string Label { get; }
        public string Origin { get; }
        public IReadOnlyDictionary<string, string> Metadata { get; }
        public string ExactText { get; }
        public bool Truncated { get; }
        public ChatReferenceNavigationState NavigationState { get; }
        public string Status { get; }
    }
}
