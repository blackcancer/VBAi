// SPDX-License-Identifier: MPL-2.0
using System;

namespace Blackcancer.ChatUI.Presentation
{
    /// <summary>Detached presentation data; positions and status are supplied by the queue owner.</summary>
    public sealed class ChatPendingMessageEntry
    {
        public ChatPendingMessageEntry(string id, string preview, int position, string status)
        {
            if (string.IsNullOrEmpty(id)) throw new ArgumentException("A pending message requires an identifier.", nameof(id));
            if (position < 0) throw new ArgumentOutOfRangeException(nameof(position));
            Id = id; Preview = preview ?? string.Empty; Position = position; Status = status ?? string.Empty;
        }
        public string Id { get; }
        public string Preview { get; }
        public int Position { get; }
        public string Status { get; }
    }
}
