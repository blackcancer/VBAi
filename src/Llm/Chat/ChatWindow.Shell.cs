using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private Border transcriptHost, promptHost;
        private Button send, changes, newChat, jumpToLatest;
        private MenuItem configure, refreshModels;
        private ComboBox providerPicker, modelPicker, effortPicker;
        private TextBlock status, promptHint;
        private WrapPanel contextChips;
        private ProgressBar activityBar;

        private static SolidColorBrush Ink(string hex)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        }

        private static Button ChatButton(string text, bool primary = false)
        {
            var button = new Button {
                Content = text, Padding = new Thickness(12, 7, 12, 7),
                Background = Ink(primary ? "#2563EB" : "#F1F5F9"),
                Foreground = Ink(primary ? "#FFFFFF" : "#334155"),
                BorderBrush = Brushes.Transparent, BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand, FontSize = 12, FontFamily = new FontFamily("Segoe UI"),
                MinHeight = 32
            };
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(7));
            border.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            border.SetBinding(Border.PaddingProperty, new System.Windows.Data.Binding("Padding") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            border.AppendChild(content);
            button.Template = new ControlTemplate(typeof(Button)) { VisualTree = border };
            var style = new Style(typeof(Button));
            var hover = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(UIElement.OpacityProperty, 0.82));
            style.Triggers.Add(hover);
            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 0.45));
            style.Triggers.Add(disabled);
            button.Style = style;
            AutomationProperties.SetName(button, text);
            return button;
        }

        private static ComboBox ChatPicker(string name)
        {
            var picker = new ComboBox { MinHeight = 30, MaxDropDownHeight = 300,
                FontSize = 12, Padding = new Thickness(7, 4, 7, 4),
                Background = Brushes.White, BorderBrush = Ink("#E2E8F0"),
                HorizontalContentAlignment = HorizontalAlignment.Stretch, ToolTip = name };
            picker.Template = (ControlTemplate)System.Windows.Markup.XamlReader.Parse(@"
<ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                 xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' TargetType='ComboBox'>
  <Grid>
    <ToggleButton Focusable='False' ClickMode='Press'
      IsChecked='{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}'>
      <ToggleButton.Template>
        <ControlTemplate TargetType='ToggleButton'>
          <Border Background='White' BorderBrush='#E2E8F0' BorderThickness='1' CornerRadius='7'>
            <TextBlock Text='⌄' HorizontalAlignment='Right' VerticalAlignment='Center' Margin='0,0,9,2' Foreground='#64748B'/>
          </Border>
        </ControlTemplate>
      </ToggleButton.Template>
    </ToggleButton>
    <ContentPresenter Content='{TemplateBinding SelectionBoxItem}' ContentTemplate='{TemplateBinding SelectionBoxItemTemplate}'
      IsHitTestVisible='False' Margin='9,5,28,5' VerticalAlignment='Center'/>
    <Popup x:Name='PART_Popup' Placement='Bottom' IsOpen='{TemplateBinding IsDropDownOpen}' AllowsTransparency='True' Focusable='False'>
      <Border Background='White' BorderBrush='#CBD5E1' BorderThickness='1' CornerRadius='7' Padding='4'
              MinWidth='{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}'>
        <ScrollViewer MaxHeight='280' CanContentScroll='True'><ItemsPresenter/></ScrollViewer>
      </Border>
    </Popup>
  </Grid>
  <ControlTemplate.Triggers>
    <Trigger Property='IsEnabled' Value='False'><Setter Property='Opacity' Value='0.45'/></Trigger>
  </ControlTemplate.Triggers>
</ControlTemplate>");
            AutomationProperties.SetName(picker, name);
            return picker;
        }

        private void InitializeShell()
        {
            var root = new Grid { Background = Ink("#F8FAFC") };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var heading = new DockPanel { Margin = new Thickness(20, 14, 20, 12), LastChildFill = true };
            var options = ChatButton("···");
            options.ToolTip = "Réglages et modèles";
            options.ContextMenu = new ContextMenu();
            configure = new MenuItem { Header = "Paramètres du fournisseur…" };
            refreshModels = new MenuItem { Header = "Actualiser les modèles" };
            options.ContextMenu.Items.Add(configure);
            options.ContextMenu.Items.Add(refreshModels);
            options.Click += (s, e) => { options.ContextMenu.PlacementTarget = options; options.ContextMenu.IsOpen = true; };
            DockPanel.SetDock(options, System.Windows.Controls.Dock.Right);
            heading.Children.Add(options);
            newChat = ChatButton("+ Nouveau");
            newChat.ToolTip = "Nouvelle conversation · Ctrl+N";
            newChat.Margin = new Thickness(0, 0, 8, 0);
            newChat.Click += (s, e) => StartNewChat();
            DockPanel.SetDock(newChat, System.Windows.Controls.Dock.Right);
            heading.Children.Add(newChat);
            var title = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            title.Children.Add(new TextBlock { Text = "CodexVBE", FontSize = 19,
                FontWeight = FontWeights.SemiBold, Foreground = Ink("#0F172A") });
            sessionTitle = new TextBlock { Text = "Votre espace de travail VBA", FontSize = 11, Foreground = Ink("#64748B"), TextTrimming = TextTrimming.CharacterEllipsis };
            title.Children.Add(sessionTitle);
            heading.Children.Add(title);
            var top = new StackPanel();
            top.Children.Add(heading);
            var scopeRow = new DockPanel { Margin = new Thickness(20, 0, 20, 10) };
            var history = ChatButton("☰ Chats");
            history.Margin = new Thickness(0, 0, 8, 0);
            history.Click += (s, e) => historyPanel.Visibility = historyPanel.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
            scopeRow.Children.Add(history);
            scopePicker = ChatPicker("Projet VBA · historique du document");
            scopeRow.Children.Add(scopePicker);
            top.Children.Add(scopeRow);
            root.Children.Add(top);
            var conversation = new Grid();
            transcriptHost = new Border();
            conversation.Children.Add(transcriptHost);
            jumpToLatest = ChatButton("↓ Dernier message");
            jumpToLatest.HorizontalAlignment = HorizontalAlignment.Center;
            jumpToLatest.VerticalAlignment = VerticalAlignment.Bottom;
            jumpToLatest.Margin = new Thickness(0, 0, 0, 10);
            jumpToLatest.Visibility = Visibility.Collapsed;
            jumpToLatest.Click += (s, e) => { followConversation = true; conversationScroll?.ScrollToEnd(); };
            conversation.Children.Add(jumpToLatest);
            var historyBody = new DockPanel { Margin = new Thickness(12) };
            var searchLabel = new TextBlock { Text = "CONVERSATIONS DU DOCUMENT", FontSize = 10,
                Foreground = Ink("#64748B"), Margin = new Thickness(0, 0, 0, 8) };
            DockPanel.SetDock(searchLabel, System.Windows.Controls.Dock.Top);
            historyBody.Children.Add(searchLabel);
            historySearch = new TextBox { Margin = new Thickness(0, 0, 0, 10), Padding = new Thickness(8), ToolTip = "Rechercher une conversation" };
            AutomationProperties.SetName(historySearch, "Rechercher une conversation");
            DockPanel.SetDock(historySearch, System.Windows.Controls.Dock.Top);
            historyBody.Children.Add(historySearch);
            var historyActions = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
            showArchived = new CheckBox { Content = "Afficher les conversations archivées", FontSize = 11, Margin = new Thickness(0, 0, 0, 10) };
            showArchived.Checked += (s, e) => RefreshHistory();
            showArchived.Unchecked += (s, e) => RefreshHistory();
            historyActions.Children.Add(showArchived);
            chatTitleEditor = new TextBox { Padding = new Thickness(7), ToolTip = "Titre de la conversation" };
            AutomationProperties.SetName(chatTitleEditor, "Titre de la conversation");
            historyActions.Children.Add(chatTitleEditor);
            var historyButtons = new WrapPanel { Margin = new Thickness(0, 6, 0, 0) };
            var rename = ChatButton("Renommer");
            rename.Click += (s, e) => RenameCurrentChat();
            historyButtons.Children.Add(rename);
            var archive = ChatButton("Archiver / réactiver");
            archive.Margin = new Thickness(5, 0, 0, 0);
            archive.Click += (s, e) => ToggleArchiveCurrentChat();
            historyButtons.Children.Add(archive);
            historyActions.Children.Add(historyButtons);
            var memory = new StackPanel();
            memory.Children.Add(new TextBlock { Text = "Notes locales propres à ce document.",
                TextWrapping = TextWrapping.Wrap, FontSize = 11, Foreground = Ink("#64748B"), Margin = new Thickness(0, 6, 0, 6) });
            memoryEditor = new TextBox { AcceptsReturn = true, TextWrapping = TextWrapping.Wrap,
                Height = 90, MaxLength = 16000, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(7) };
            AutomationProperties.SetName(memoryEditor, "Mémoire du document");
            memory.Children.Add(memoryEditor);
            var saveMemory = ChatButton("Enregistrer localement");
            saveMemory.Margin = new Thickness(0, 6, 0, 0);
            saveMemory.Click += (s, e) => SaveProjectMemory();
            memory.Children.Add(saveMemory);
            attachMemory = new CheckBox { Content = new TextBlock {
                Text = "Joindre les notes enregistrées au prochain message envoyé au fournisseur sélectionné", TextWrapping = TextWrapping.Wrap },
                IsChecked = false, FontSize = 11, Margin = new Thickness(0, 8, 0, 0) };
            attachMemory.Checked += (s, e) => RefreshContextChips();
            attachMemory.Unchecked += (s, e) => RefreshContextChips();
            memory.Children.Add(attachMemory);
            historyActions.Children.Add(new Expander { Header = "Mémoire du document", Content = memory,
                Margin = new Thickness(0, 12, 0, 0), FontSize = 12 });
            DockPanel.SetDock(historyActions, System.Windows.Controls.Dock.Bottom);
            historyBody.Children.Add(historyActions);
            sessionList = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent,
                FontSize = 13, HorizontalContentAlignment = HorizontalAlignment.Stretch };
            var sessionStyle = new Style(typeof(ListBoxItem));
            sessionStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 10, 8, 10)));
            sessionList.ItemContainerStyle = sessionStyle;
            historyBody.Children.Add(sessionList);
            historyPanel = new Border { Child = historyBody, Width = 290, HorizontalAlignment = HorizontalAlignment.Left,
                Background = Brushes.White, BorderBrush = Ink("#CBD5E1"), BorderThickness = new Thickness(0, 1, 1, 1),
                Visibility = Visibility.Collapsed };
            conversation.Children.Add(historyPanel);
            Grid.SetRow(conversation, 1);
            root.Children.Add(conversation);

            var bottom = new StackPanel { Margin = new Thickness(18, 8, 18, 12) };
            var composer = new StackPanel();
            contextChips = new WrapPanel { Margin = new Thickness(10, 8, 10, 0), Visibility = Visibility.Collapsed };
            composer.Children.Add(contextChips);
            var entry = new Grid { Margin = new Thickness(4) };
            promptHost = new Border { MinHeight = 82, MaxHeight = 160 };
            entry.Children.Add(promptHost);
            promptHint = new TextBlock { Text = "Posez une question ou décrivez une modification…",
                Foreground = Ink("#94A3B8"), Margin = new Thickness(12, 12, 12, 0),
                IsHitTestVisible = false, TextWrapping = TextWrapping.Wrap, FontSize = 13 };
            entry.Children.Add(promptHint);
            composer.Children.Add(entry);
            var actions = new DockPanel { Margin = new Thickness(10, 0, 10, 10) };
            send = ChatButton("Envoyer ↑", true);
            DockPanel.SetDock(send, System.Windows.Controls.Dock.Right);
            actions.Children.Add(send);
            var references = new WrapPanel();
            var modules = ChatButton("# Contexte");
            modules.ToolTip = "Référencer un projet ou un module";
            modules.Click += (s, e) => InsertReferencePrefix('#');
            references.Children.Add(modules);
            var methods = ChatButton("@ Fonction");
            methods.Margin = new Thickness(6, 0, 0, 0);
            methods.ToolTip = "Référencer une Sub, Function ou Property";
            methods.Click += (s, e) => InsertReferencePrefix('@');
            references.Children.Add(methods);
            actions.Children.Add(references);
            composer.Children.Add(actions);
            bottom.Children.Add(new Border { Child = composer, CornerRadius = new CornerRadius(12),
                BorderBrush = Ink("#CBD5E1"), BorderThickness = new Thickness(1), Background = Brushes.White });

            var pickers = new Grid { Margin = new Thickness(0, 8, 0, 6) };
            pickers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.85, GridUnitType.Star) });
            pickers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.4, GridUnitType.Star) });
            pickers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.85, GridUnitType.Star) });
            providerPicker = ChatPicker("Fournisseur");
            modelPicker = ChatPicker("Modèle");
            effortPicker = ChatPicker("Raisonnement");
            effortPicker.IsEnabled = false;
            modelPicker.Margin = new Thickness(6, 0, 6, 0);
            Grid.SetColumn(modelPicker, 1);
            Grid.SetColumn(effortPicker, 2);
            pickers.Children.Add(providerPicker); pickers.Children.Add(modelPicker); pickers.Children.Add(effortPicker);
            bottom.Children.Add(pickers);
            var toolsRow = new DockPanel();
            changes = ChatButton("Modifications · 0");
            changes.IsEnabled = false;
            changes.Padding = new Thickness(8, 4, 8, 4);
            DockPanel.SetDock(changes, System.Windows.Controls.Dock.Right); toolsRow.Children.Add(changes);
            toolsRow.Children.Add(new TextBlock { Text = "Entrée : envoyer · Maj+Entrée : nouvelle ligne",
                FontSize = 10, Foreground = Ink("#64748B"), VerticalAlignment = VerticalAlignment.Center });
            bottom.Children.Add(toolsRow);
            activityBar = new ProgressBar { Height = 2, IsIndeterminate = true,
                Foreground = Ink("#2563EB"), Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) };
            bottom.Children.Add(activityBar);
            status = new TextBlock { Text = "Connexion…", FontSize = 11,
                TextTrimming = TextTrimming.CharacterEllipsis, Foreground = Ink("#64748B"), Margin = new Thickness(0, 6, 0, 0) };
            bottom.Children.Add(status);
            Grid.SetRow(bottom, 2); root.Children.Add(bottom);
            root.PreviewKeyDown += (s, e) => {
                if (e.Key == Key.N && Keyboard.Modifiers == System.Windows.Input.ModifierKeys.Control && !busy)
                { StartNewChat(); e.Handled = true; }
            };
            shellHost.Child = root;
        }

        private void InsertReferencePrefix(char prefix)
        {
            if (prompt == null) return;
            prompt.Focus();
            int caret = prompt.CaretIndex;
            prompt.SelectedText = (caret > 0 && !char.IsWhiteSpace(prompt.Text[caret - 1]) ? " " : "") + prefix;
            prompt.CaretIndex = prompt.SelectionStart + prompt.SelectionLength;
            UpdateReferences();
        }

        private void StartNewChat()
        {
            NewSession();
            prompt?.Focus();
        }
    }
}
