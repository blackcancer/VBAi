// SPDX-License-Identifier: MPL-2.0
using System;

namespace Blackcancer.ChatUI.Presentation
{
    /// <summary>An inert host-defined reading choice. Identifiers carry no execution permission.</summary>
    public sealed class ChatReadingChoice
    {
        public ChatReadingChoice(string id, string label)
        {
            if (string.IsNullOrEmpty(id)) throw new ArgumentException("A reading choice needs an identifier.", nameof(id));
            Id = id;
            Label = label ?? string.Empty;
        }
        public string Id { get; }
        public string Label { get; }
        public override string ToString() => Label;
    }
}
