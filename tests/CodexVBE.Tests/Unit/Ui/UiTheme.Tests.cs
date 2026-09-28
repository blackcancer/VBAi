using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie la sélection, la propagation et la présentation des thèmes de l’interface.</summary>
    [TestClass]
    public sealed class UiThemeTests
    {
        /// <summary>Choisit la palette selon préférences système, contraste et thème explicitement sélectionné.</summary>
        [TestMethod]
        public void SystemPreferencesContrastAndExplicitThemesSelectExpectedPalette()
        {
            Assert.AreEqual(SystemInformation.HighContrast,UiTheme.HighContrast());
            Assert.AreEqual(SystemColors.Window,UiTheme.WindowColor());
            UiTheme.ReadSystemTheme();
            using(var scope=new ThemeScope())
            {
                ThemeScope.SetChoice(ThemeChoice.System);
                foreach(var value in new object[]{0,1,null}) { UiTheme.ReadSystemTheme=()=>value; Assert.AreEqual(Equals(value,0),UiTheme.Dark); }
                UiTheme.ReadSystemTheme=()=>{throw new IOException("registry unavailable");}; Assert.IsFalse(UiTheme.Dark);
                UiTheme.HighContrast=()=>true;
                UiTheme.WindowColor=()=>Color.Black; Assert.IsTrue(UiTheme.Dark);
                UiTheme.WindowColor=()=>Color.White; Assert.IsFalse(UiTheme.Dark);
                Assert.AreEqual(SystemColors.Window,UiTheme.Surface); Assert.AreEqual(SystemColors.Control,UiTheme.Background); Assert.AreEqual(SystemColors.WindowText,UiTheme.Foreground);
                UiTheme.HighContrast=()=>false;
                foreach(var dark in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    Assert.AreEqual(dark,UiTheme.Dark);
                    Assert.AreEqual(dark ? Color.FromArgb(30,34,42) : Color.White,UiTheme.Surface);
                    Assert.AreEqual(dark ? Color.FromArgb(22,26,33) : Color.FromArgb(248,250,252),UiTheme.Background);
                    Assert.AreEqual(dark ? Color.FromArgb(226,232,240) : Color.FromArgb(30,41,59),UiTheme.Foreground);
                    Assert.AreEqual(dark ? Color.FromArgb(24,64,42) : Color.FromArgb(232,247,237),UiTheme.Added);
                    Assert.AreEqual(dark ? Color.FromArgb(78,35,40) : Color.FromArgb(255,240,240),UiTheme.Removed);
                }
            }
        }
        /// <summary>Charge et sauvegarde le thème et notifie les abonnés après validation.</summary>
        [TestMethod]
        public void LoadingAndSavingThemeValidateStoredEnumAndNotifySubscribers()
        {
            using(var scope=new ThemeScope())
            {
                foreach(var text in new[]{"Light","Dark","System","99","invalid"})
                {
                    UiTheme.ReadTheme=p=>text;
                    var actual=(ThemeChoice)UiInvoke.Call(typeof(UiTheme),"Load",null);
                    Assert.AreEqual(text=="Light" ? ThemeChoice.Light : text=="Dark" ? ThemeChoice.Dark : ThemeChoice.System,actual);
                }
                UiTheme.ReadTheme=p=>{throw new IOException("missing");}; Assert.AreEqual(ThemeChoice.System,UiInvoke.Call(typeof(UiTheme),"Load",null));
                string path=Path.Combine(Path.GetTempPath(),"CodexVBE-theme-"+Guid.NewGuid().ToString("N"),"theme.txt");
                UiTheme.FileName=path; UiTheme.WriteTheme=File.WriteAllText;
                Action changed=null;
                try
                {
                    UiTheme.Select(ThemeChoice.Light); Assert.AreEqual("Light",File.ReadAllText(path)); Assert.AreEqual(ThemeChoice.Light,UiTheme.Choice);
                    UiInvoke.Call(typeof(UiTheme),"PreferencesChanged",null,null,new UserPreferenceChangedEventArgs(UserPreferenceCategory.General));
                    int count=0; changed=()=>count++; UiTheme.Changed+=changed;
                    UiTheme.Select(ThemeChoice.Dark); Assert.AreEqual("Dark",File.ReadAllText(path)); Assert.AreEqual(1,count);
                    UiInvoke.Call(typeof(UiTheme),"PreferencesChanged",null,null,new UserPreferenceChangedEventArgs(UserPreferenceCategory.Color)); Assert.AreEqual(2,count);
                }
                finally { if(changed!=null) UiTheme.Changed-=changed; if(File.Exists(path)) File.Delete(path); Directory.Delete(Path.GetDirectoryName(path)); }
            }
        }
        /// <summary>Configure récursivement les contrôles et fenêtres natives lors de l’application d’un thème.</summary>
        [STATestMethod]
        public void ApplyingThemeConfiguresControlsRecursivelyAndNativeHandles()
        {
            using(var scope=new ThemeScope())
            using(var form=new Form())
            using(var text=new TextBox())
            using(var list=new ListBox())
            using(var check=new CheckBox())
            using(var button=new Button())
            using(var combo=new ComboBox())
            using(var grid=new DataGridView())
            {
                form.Controls.AddRange(new Control[]{text,list,check,button,combo,grid});
                foreach(var dark in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    UiTheme.Apply(form); // pre-handle registration
                    foreach(Control control in form.Controls) { var handle=control.Handle; Assert.AreNotEqual(IntPtr.Zero,handle); }
                    var formHandle=form.Handle; UiTheme.Apply(form); // handles already created
                    Assert.AreEqual(UiTheme.Background,form.BackColor); Assert.AreEqual(UiTheme.Foreground,form.ForeColor);
                    Assert.AreEqual(UiTheme.Surface,text.BackColor); Assert.AreEqual(BorderStyle.FixedSingle,text.BorderStyle);
                    Assert.AreEqual(BorderStyle.FixedSingle,list.BorderStyle); Assert.AreEqual(FlatStyle.Flat,check.FlatStyle);
                    Assert.AreEqual(FlatStyle.Flat,button.FlatStyle); Assert.AreEqual(dark ? Color.FromArgb(75,85,99) : Color.FromArgb(203,213,225),button.FlatAppearance.BorderColor);
                    Assert.AreEqual(DrawMode.OwnerDrawFixed,combo.DrawMode); Assert.AreEqual(FlatStyle.Flat,combo.FlatStyle);
                    Assert.IsFalse(grid.EnableHeadersVisualStyles); Assert.AreEqual(UiTheme.Surface,grid.DefaultCellStyle.BackColor); Assert.AreEqual(UiTheme.Foreground,grid.DefaultCellStyle.ForeColor);
                    Assert.AreEqual(UiTheme.Background,grid.ColumnHeadersDefaultCellStyle.SelectionBackColor); Assert.AreEqual(UiTheme.Foreground,grid.ColumnHeadersDefaultCellStyle.SelectionForeColor);
                    Assert.AreEqual(dark ? Color.FromArgb(71,85,105) : Color.FromArgb(203,213,225),grid.GridColor);
                }
                UiTheme.HighContrast=()=>true; UiTheme.Apply(form);
                Assert.AreEqual(SystemColors.Control,form.BackColor);
                using(var uncreated=new Control()) UiInvoke.Call(typeof(UiTheme),"ApplyNativeTheme",null,uncreated,EventArgs.Empty);
                var context=LicenseManager.CurrentContext;
                try { LicenseManager.CurrentContext=new DesignContext(); UiTheme.Attach(form); UiInvoke.Call(typeof(UiTheme),"ApplyNativeTheme",null,form,EventArgs.Empty); Assert.AreEqual(context.UsageMode,LicenseUsageMode.Runtime); }
                finally { LicenseManager.CurrentContext=context; }
            }
        }
        /// <summary>Met à jour les formulaires attachés sur les deux threads et se désabonne à leur disposal.</summary>
        [STATestMethod]
        public void AttachedFormsUpdateOnBothThreadsAndUnsubscribeWhenDisposed()
        {
            using(var scope=new ThemeScope())
            {
                var initial=(Action)typeof(UiTheme).GetField("Changed",BindingFlags.Static|BindingFlags.NonPublic).GetValue(null);
                var form=new Form(); UiTheme.Attach(form); var handle=form.Handle;
                ThemeScope.SetChoice(ThemeChoice.Dark); UiTheme.Select(ThemeChoice.Dark); Assert.AreEqual(UiTheme.Background,form.BackColor);
                ThemeScope.SetChoice(ThemeChoice.Light);
                var thread=new Thread(()=>UiTheme.Select(ThemeChoice.Light)); thread.Start(); Assert.IsTrue(thread.Join(5000)); Application.DoEvents(); Assert.AreEqual(UiTheme.Background,form.BackColor);
                var handler=(Action)typeof(UiTheme).GetField("Changed",BindingFlags.Static|BindingFlags.NonPublic).GetValue(null);
                form.Dispose(); handler();
                Assert.AreEqual(initial,typeof(UiTheme).GetField("Changed",BindingFlags.Static|BindingFlags.NonPublic).GetValue(null));
            }
        }
        /// <summary>Formate les cellules de diff en ignorant les lignes ordinaires et virtuelles invalides.</summary>
        [STATestMethod]
        public void DiffFormattingIgnoresVirtualInvalidAndOrdinaryCellsAndColorsBothChangeKinds()
        {
            using(var scope=new ThemeScope())
            using(var grid=new DataGridView {AllowUserToAddRows=false})
            {
                grid.Columns.Add("code","Code"); grid.Rows.Add("text");
                foreach(var tuple in new[]{new[]{-1,0},new[]{0,-1},new[]{1,0},new[]{0,0}})
                {
                    var args=new DataGridViewCellFormattingEventArgs(tuple[1],tuple[0],"text",typeof(string),new DataGridViewCellStyle{BackColor=Color.Magenta});
                    UiInvoke.Call(typeof(UiTheme),"FormatDiffCell",null,grid,args); Assert.AreEqual(Color.Magenta,args.CellStyle.BackColor);
                }
                grid.VirtualMode=true;
                var virtualArgs=new DataGridViewCellFormattingEventArgs(0,0,"text",typeof(string),new DataGridViewCellStyle{BackColor=Color.Magenta});
                UiInvoke.Call(typeof(UiTheme),"FormatDiffCell",null,grid,virtualArgs); Assert.AreEqual(Color.Magenta,virtualArgs.CellStyle.BackColor); grid.VirtualMode=false;
                foreach(var kind in new[]{"vba-added","vba-removed"})
                {
                    grid.Rows[0].Cells[0].Tag=kind;
                    var args=new DataGridViewCellFormattingEventArgs(0,0,"text",typeof(string),new DataGridViewCellStyle());
                    UiInvoke.Call(typeof(UiTheme),"FormatDiffCell",null,grid,args);
                    Assert.AreEqual(kind=="vba-added" ? UiTheme.Added : UiTheme.Removed,args.CellStyle.BackColor); Assert.AreEqual(UiTheme.Foreground,args.CellStyle.ForeColor);
                }
            }
        }
        /// <summary>Dessine les choix de ComboBox avec les couleurs de sélection et le texte de repli.</summary>
        [STATestMethod]
        public void OwnerDrawComboUsesSelectionColorsAndFallbackText()
        {
            using(var scope=new ThemeScope())
            using(var combo=new ComboBox {Text="fallback"})
            using(var bitmap=new Bitmap(200,30))
            using(var graphics=Graphics.FromImage(bitmap))
            {
                combo.Items.Add("item"); UiTheme.Apply(combo); UiTheme.Apply(combo);
                foreach(var index in new[]{-1,0})
                foreach(var selected in new[]{false,true})
                {
                    var args=new DrawItemEventArgs(graphics,combo.Font,new Rectangle(0,0,200,30),index,selected ? DrawItemState.Selected | DrawItemState.Focus : DrawItemState.None);
                    UiInvoke.Call(typeof(UiTheme),"DrawCombo",null,combo,args);
                    Assert.AreEqual((selected ? SystemColors.Highlight : UiTheme.Surface).ToArgb(),bitmap.GetPixel(190,15).ToArgb());
                }
            }
        }
        /// <summary>Préserve les couleurs claires ou inconnues et convertit chaque entrée de la palette sombre.</summary>
        [TestMethod]
        public void MappingPreservesLightAndUnknownColorsAndMapsEveryDarkPaletteEntry()
        {
            using(var scope=new ThemeScope())
            {
                Assert.AreEqual("#FFFFFF",UiTheme.Map("#FFFFFF")); ThemeScope.SetChoice(ThemeChoice.Dark);
                var groups=new[]{"#FFFFFF=#1E222A","#F8FAFC=#161A21","#F1F5F9=#273345","#EFF6FF=#273345","#DBEAFE=#526176","#E2E8F0=#526176","#0F172A=#E2E8F0","#1E293B=#E2E8F0","#334155=#E2E8F0","#475569=#E2E8F0","#64748B=#AEBBD0","#FEF2F2=#4E2328","#FFF0F0=#4E2328","#E8F7ED=#18402A","#15803D=#86EFAC","#166534=#86EFAC","#991B1B=#FCA5A5"};
                foreach(var group in groups) {var parts=group.Split('='); Assert.AreEqual(parts[1],UiTheme.Map(parts[0].ToLowerInvariant()));}
                Assert.AreEqual("#123456",UiTheme.Map("#123456"));
            }
        }

    }
}
