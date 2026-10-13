using System;

namespace Blackcancer.ChatUI.CommonRuntime
{
    /// <summary>Detached presentation values shared by host activity adapters.</summary>
    public sealed class ActivityRenderSnapshot
    {
        public ActivityRenderSnapshot(string kind, string title, string detail, string status,
            long? durationMilliseconds, string cultureName, bool expanded)
        {
            Kind = kind;
            Title = title;
            Detail = detail;
            Status = status;
            DurationMilliseconds = durationMilliseconds;
            CultureName = cultureName;
            Expanded = expanded;
        }

        public string Kind { get; }
        public string Title { get; }
        public string Detail { get; }
        public string Status { get; }
        public long? DurationMilliseconds { get; }
        public string CultureName { get; }
        public bool Expanded { get; }

        internal bool SamePresentation(ActivityRenderSnapshot other) => other != null &&
            string.Equals(Kind, other.Kind, StringComparison.Ordinal) &&
            string.Equals(Title, other.Title, StringComparison.Ordinal) &&
            string.Equals(Detail, other.Detail, StringComparison.Ordinal) &&
            string.Equals(Status, other.Status, StringComparison.Ordinal) &&
            DurationMilliseconds == other.DurationMilliseconds &&
            string.Equals(CultureName, other.CultureName, StringComparison.Ordinal) &&
            Expanded == other.Expanded;
    }

    /// <summary>Caches only successfully rendered values; no provider, UI or host objects are retained.</summary>
    public sealed class ActivityRenderState
    {
        private ActivityRenderSnapshot rendered;

        public bool IsCurrent(ActivityRenderSnapshot presentation) =>
            rendered != null && rendered.SamePresentation(presentation);

        /// <summary>Publish after updating all controls so a failed update remains retryable.</summary>
        public void RecordRendered(ActivityRenderSnapshot presentation)
        {
            rendered = presentation ?? throw new ArgumentNullException(nameof(presentation));
        }
    }
}
