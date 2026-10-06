using System;
using System.Windows;
using System.Windows.Controls;

namespace VBAi
{
    // The item container is recycled; expensive message views exist only while realized.
    /// <summary>Conteneur WPF recyclé qui crée son contenu à l’affichage et le libère lorsqu’il quitte l’écran.</summary>
    public sealed class TranscriptItem : ContentControl
    {

        /// <summary>Propriété de dépendance du callback qui crée le contenu d’un élément réalisé.</summary>
        public static readonly DependencyProperty RenderProperty = DependencyProperty.Register("Render", typeof(Action<TranscriptItem>), typeof(TranscriptItem));

        /// <summary>Propriété de dépendance du callback qui libère le contenu d’un élément recyclé.</summary>
        public static readonly DependencyProperty ReleaseProperty = DependencyProperty.Register("Release", typeof(Action<TranscriptItem>), typeof(TranscriptItem));

        /// <summary>Contexte de données pour lequel le contenu actuellement réalisé a été créé.</summary>
        internal object RenderedContext;

        /// <summary>Configure l’étirement du contenu et relie le cycle de vie WPF à la réalisation ou libération du contenu.</summary>
        public TranscriptItem()
        {
            HorizontalContentAlignment = HorizontalAlignment.Stretch;
            Loaded += (s, e) => Realize();
            Unloaded += (s, e) => Clear();
            DataContextChanged += (s, e) => { Clear(); if (IsLoaded) Realize(); };
        }

        /// <summary>Crée le contenu une seule fois pour le contexte actuel, si le conteneur est vide.</summary>
        private void Realize()
        {
            if (Content != null) return;
            RenderedContext = DataContext;
            ((Action<TranscriptItem>)GetValue(RenderProperty))?.Invoke(this);
        }

        /// <summary>Libère le contenu réalisé puis efface le contenu et son contexte.</summary>
        private void Clear()
        {
            if (Content == null) return;
            ((Action<TranscriptItem>)GetValue(ReleaseProperty))?.Invoke(this);
            Content = null; RenderedContext = null;
        }
    }
}
