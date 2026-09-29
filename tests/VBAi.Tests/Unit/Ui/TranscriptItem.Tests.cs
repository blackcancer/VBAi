using System;
using System.Windows;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie le recyclage et le rendu des éléments de transcription.</summary>
    [TestClass]
    public sealed class TranscriptItemTests
    {
        /// <summary>Libère l’ancienne vue lors du recyclage et évite un second rendu.</summary>
        [STATestMethod]
        public void RecyclingReleasesOldViewAndDoesNotRenderTwice()
        {
            var item=new TranscriptItem(); int renders=0,releases=0;
            item.DataContext="first";
            item.SetValue(TranscriptItem.RenderProperty,new Action<TranscriptItem>(i=>{renders++; i.Content="view";}));
            item.SetValue(TranscriptItem.ReleaseProperty,new Action<TranscriptItem>(i=>{Assert.AreEqual("first",i.RenderedContext); releases++;}));
            item.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
            item.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
            Assert.AreEqual(1,renders); Assert.AreEqual("first",item.RenderedContext);
            item.DataContext="second";
            Assert.AreEqual(1,releases); Assert.IsNull(item.Content); Assert.IsNull(item.RenderedContext);
            item.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent));
            Assert.AreEqual(1,releases);
            item.ClearValue(TranscriptItem.RenderProperty); item.ClearValue(TranscriptItem.ReleaseProperty);
            item.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
            Assert.IsNull(item.Content); Assert.AreEqual("second",item.RenderedContext);
            item.Content="external";
            item.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent));
            Assert.IsNull(item.Content); Assert.IsNull(item.RenderedContext);
        }
        /// <summary>Rend le contexte de remplacement après chargement du conteneur.</summary>
        [STATestMethod]
        public void LoadedContainerRendersReplacementContext()
        {
            var item=new TranscriptItem(); int renders=0,releases=0;
            item.SetValue(TranscriptItem.RenderProperty,new Action<TranscriptItem>(i=>{renders++;i.Content=i.DataContext;}));
            item.SetValue(TranscriptItem.ReleaseProperty,new Action<TranscriptItem>(i=>releases++));
            var window=new Window { Content=item,Width=1,Height=1,ShowInTaskbar=false,WindowStyle=WindowStyle.None,Left=-10000,Top=-10000 };
            try { item.DataContext="one"; window.Show(); item.DataContext="two"; Assert.IsTrue(item.IsLoaded); Assert.AreEqual("two",item.Content); Assert.AreEqual(2,renders); Assert.AreEqual(1,releases); }
            finally { window.Close(); }
        }
    }
}
