using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class ChatDiffViewTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void NavigationSearchExpansionAndLayoutOptionsPreserveChanges()
        {
            using(var theme=new ThemeScope())
            using(var culture=new LocalizationScope())
            {
                var before=string.Join("\n",Enumerable.Range(0,25).Select(i=>"line "+i)); var after=before.Replace("line 5\n","changed five\n").Replace("line 20\n","changed twenty\n");
                var view=new ChatDiffView(before,after); var grid=UiInvoke.Field<DataGrid>(view,"grid"); var unified=UiInvoke.Field<CheckBox>(view,"unified"); var fold=UiInvoke.Field<CheckBox>(view,"fold"); var search=UiInvoke.Field<TextBox>(view,"search");
                Assert.AreEqual(3,grid.Columns.Count); Assert.AreEqual(FlowDirection.LeftToRight,view.FlowDirection);
                var buttons=((WrapPanel)view.Children[0]).Children.OfType<Button>().ToArray();
                buttons[1].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); var first=(DiffRow)grid.SelectedItem; Assert.IsTrue(first.Hunk>=0);
                buttons[1].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.AreNotEqual(first.Hunk,((DiffRow)grid.SelectedItem).Hunk);
                buttons[0].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.AreEqual(first.Hunk,((DiffRow)grid.SelectedItem).Hunk);
                grid.SelectedItem=grid.Items.OfType<DiffRow>().First(r=>r.Fold); grid.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice,0,MouseButton.Left){RoutedEvent=Control.MouseDoubleClickEvent}); Assert.IsFalse(fold.IsChecked.Value);
                grid.SelectedItem=grid.Items.OfType<DiffRow>().First(); grid.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice,0,MouseButton.Left){RoutedEvent=Control.MouseDoubleClickEvent}); grid.SelectedIndex=-1; grid.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice,0,MouseButton.Left){RoutedEvent=Control.MouseDoubleClickEvent});
                unified.IsChecked=false; unified.RaiseEvent(new RoutedEventArgs(CheckBox.ClickEvent)); Assert.AreEqual(4,grid.Columns.Count);
                fold.IsChecked=true; fold.RaiseEvent(new RoutedEventArgs(CheckBox.ClickEvent));
                buttons[2].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.IsTrue(fold.IsChecked.Value);
                search.Text="CHANGED TWENTY"; buttons[2].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.IsFalse(fold.IsChecked.Value); Assert.AreEqual("changed twenty",((DiffRow)grid.SelectedItem).Right);
                search.Text="absent"; var selected=grid.SelectedItem; buttons[2].RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.AreSame(selected,grid.SelectedItem);
                var window=new Window{Content=view,Width=1,Height=1,Left=-10000,Top=-10000,ShowInTaskbar=false};
                try {window.Show(); var source=PresentationSource.FromVisual(search); foreach(var key in new[]{Key.Escape,Key.Enter}) {var args=new KeyEventArgs(Keyboard.PrimaryDevice,source,0,key){RoutedEvent=Keyboard.KeyDownEvent}; search.RaiseEvent(args); Assert.AreEqual(key==Key.Enter,args.Handled);} }
                finally {window.Close();}
                foreach(var text in new[]{"","unchanged"}) {var empty=new ChatDiffView(text,text); UiInvoke.Call(typeof(ChatDiffView),"Move",empty,1); Assert.AreEqual(-1,UiInvoke.Field<DataGrid>(empty,"grid").SelectedIndex);}
                var one=new ChatDiffView("old","new"); UiInvoke.Call(typeof(ChatDiffView),"Move",one,1); var single=UiInvoke.Field<DataGrid>(one,"grid"); var prior=single.SelectedItem; UiInvoke.Call(typeof(ChatDiffView),"Move",one,1); Assert.AreSame(prior,single.SelectedItem);
            }
        }
        [STATestMethod,TestCategory("Unit")]
        public void SyntaxConverterColorsEachSideAndHandlesNullRowsAndMissingSources()
        {
            using(var theme=new ThemeScope())
            {
                var type=typeof(ChatDiffView).GetNestedType("SyntaxConverter",BindingFlags.NonPublic);
                foreach(var field in new[]{"Left","Right","Unified"})
                {
                    var converter=(IValueConverter)Activator.CreateInstance(type,BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{field},CultureInfo.InvariantCulture);
                    Assert.IsNull(converter.Convert(null,typeof(object),null,CultureInfo.InvariantCulture));
                    foreach(var changed in new[]{false,true})
                    foreach(var hasNew in new[]{false,true})
                    {
                        var row=new DiffRow{Left="Sub old",Right=hasNew ? "Sub new" : null,New=hasNew ? (int?)1 : null,Hunk=changed ? 0 : -1};
                        var text=(TextBlock)converter.Convert(row,typeof(object),null,CultureInfo.InvariantCulture);
                        var expected=!changed ? UiTheme.Surface : field=="Left" || field=="Unified"&&!hasNew ? UiTheme.Removed : UiTheme.Added;
                        Assert.AreEqual(Color.FromRgb(expected.R,expected.G,expected.B),((SolidColorBrush)text.Background).Color);
                        Assert.AreEqual(field=="Left" ? "Sub old" : field=="Right" ? row.Right??"" : row.Unified,string.Concat(text.Inlines.OfType<System.Windows.Documents.Run>().Select(r=>r.Text)));
                    }
                    Assert.ThrowsException<NotSupportedException>(()=>converter.ConvertBack(null,typeof(object),null,CultureInfo.InvariantCulture));
                }
            }
        }
    }
}
