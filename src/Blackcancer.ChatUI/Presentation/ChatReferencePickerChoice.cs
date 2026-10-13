// SPDX-License-Identifier: MPL-2.0
using System;

namespace Blackcancer.ChatUI.Presentation
{
    /// <summary>Display data and an opaque detached host witness. The UI never interprets the witness.</summary>
    public sealed class ChatReferencePickerChoice
    {
        public ChatReferencePickerChoice(string id, string label, string kind, string detail, string detachedWitness)
        {
            if (string.IsNullOrEmpty(id)) throw new ArgumentException("A choice requires an identifier.", nameof(id));
            Id = id; Label = label ?? string.Empty; Kind = kind ?? string.Empty;
            Detail = detail ?? string.Empty; DetachedWitness = detachedWitness ?? string.Empty;
        }
        public string Id { get; }
        public string Label { get; }
        public string Kind { get; }
        public string Detail { get; }
        public string DetachedWitness { get; }
    }
}
