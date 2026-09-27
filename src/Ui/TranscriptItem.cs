using System;
using System.Windows;
using System.Windows.Controls;

namespace CodexVBE
{
    // The item container is recycled; expensive message views exist only while realized.
    public sealed class TranscriptItem : ContentControl
    {
        public static readonly DependencyProperty RenderProperty = DependencyProperty.Register("Render", typeof(Action<TranscriptItem>), typeof(TranscriptItem));
        public static readonly DependencyProperty ReleaseProperty = DependencyProperty.Register("Release", typeof(Action<TranscriptItem>), typeof(TranscriptItem));
        internal object RenderedContext;
        public TranscriptItem()
        {
            HorizontalContentAlignment = HorizontalAlignment.Stretch;
            Loaded += (s, e) => Realize();
            Unloaded += (s, e) => Clear();
            DataContextChanged += (s, e) => { Clear(); if (IsLoaded) Realize(); };
        }
        private void Realize()
        {
            if (Content != null) return;
            RenderedContext = DataContext;
            ((Action<TranscriptItem>)GetValue(RenderProperty))?.Invoke(this);
        }
        private void Clear()
        {
            if (Content == null) return;
            ((Action<TranscriptItem>)GetValue(ReleaseProperty))?.Invoke(this);
            Content = null; RenderedContext = null;
        }
    }
}
