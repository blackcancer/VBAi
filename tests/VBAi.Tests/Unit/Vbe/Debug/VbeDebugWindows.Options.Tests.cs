using System.Windows.Automation;

using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed class WritableOptionsTests
    {
        /// <summary>La frontière injectée envoie une seule notification de changement au parent exact; aucune preuve native n'est simulée.</summary>
        [TestMethod]
        public void FormatCategorySelectionNotifiesItsQualifiedParentExactlyOnceAndPropagatesFailures()
        {
            var list = new IntPtr(71); var parent = new IntPtr(19); int calls = 0;
            VbeDebugWindows.NotifyOptionsListSelection(list, 4905, parent, (window, message, argument, value) =>
            {
                calls++; Assert.AreEqual(parent, window); Assert.AreEqual(0x111, message);
                Assert.AreEqual(4905L, argument.ToInt64() & 0xffff); Assert.AreEqual(1L, argument.ToInt64() >> 16);
                Assert.AreEqual(list, value);
            });
            Assert.AreEqual(1, calls);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.NotifyOptionsListSelection(list, 4905, parent,
                (window, message, argument, value) => { calls++; throw new InvalidOperationException("notification failed"); }));
            Assert.AreEqual(2, calls, "A failed notification is never resent automatically.");
            foreach (int scenario in Enumerable.Range(0, 6))
            {
                var child = scenario == 0 ? IntPtr.Zero : list;
                var owner = scenario == 1 ? IntPtr.Zero : scenario == 2 ? list : parent;
                int id = scenario == 3 ? -1 : scenario == 4 ? 65536 : 4905;
                Action<IntPtr, int, IntPtr, IntPtr> dispatch = scenario == 5 ? null : (Action<IntPtr, int, IntPtr, IntPtr>)((window, message, argument, value) => calls++);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.NotifyOptionsListSelection(child, id, owner, dispatch));
            }
            Assert.AreEqual(2, calls);
        }

        /// <summary>L'inspection de dix catégories appelle uniquement la lecture palette une fois chacune puis restaure la sélection.</summary>
        [TestMethod]
        public void FormatCategoryInspectionReadsOnlyEachPaletteSetAndRestoresOriginalSelection()
        {
            var labels = Enumerable.Range(0, 10).Select(i => "Category " + i).ToArray();
            var list = new VbeDebugWindows.OptionsControl { Choices = labels, Value = labels[4] };
            var selected = new List<string>(); int reads = 0;
            var categories = VbeDebugWindows.CaptureOptionsFormatCategories(list, name => selected.Add(name), () =>
            {
                reads++; return new[] { "Foreground", "Background", "Indicator" }.Select(name =>
                    new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.ComboBox", Value = "NativeIndex:1", Enabled = name != "Indicator" }).ToArray();
            });
            Assert.AreEqual(10, reads); Assert.AreEqual(10, categories.Count);
            CollectionAssert.AreEqual(labels.Concat(new[] { labels[4] }).ToArray(), selected.ToArray());
            Assert.IsFalse(categories[0].Palettes[2].Enabled);
            Assert.AreEqual("NativeIndex:1", categories[9].Palettes[0].Value);
        }

        /// <summary>Les échecs de sélection ou lecture et tous les contrôles invalides restaurent l'état avant propagation.</summary>
        [TestMethod]
        public void FormatCategoryInspectionRestoresOnEverySelectionAndPaletteFailure()
        {
            foreach (int scenario in Enumerable.Range(0, 10))
            {
                var list = new VbeDebugWindows.OptionsControl { Choices = new[] { "Original", "Other" }, Value = "Original" };
                var selected = new List<string>();
                var palettes = new[] { "Foreground", "Background", "Indicator" }.Select(name =>
                    new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.ComboBox", Value = "Automatic" }).ToArray();
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CaptureOptionsFormatCategories(list, name =>
                {
                    selected.Add(name); if (scenario == 0 && name == "Other") throw new InvalidOperationException("selection failed");
                }, () =>
                {
                    if (scenario == 1) throw new InvalidOperationException("palette failed");
                    if (scenario == 2) return null;
                    if (scenario == 3) return palettes.Take(2).ToArray();
                    if (scenario == 4) palettes[0].Type = "ControlType.Edit";
                    if (scenario == 5) palettes[0].Name = "Font";
                    if (scenario == 6) palettes[0].Name = "Background";
                    if (scenario == 7) palettes[0].Error = "unreadable";
                    if (scenario == 8) palettes[0].Value = null;
                    if (scenario == 9) palettes[0].Visible = false;
                    return palettes;
                }));
                Assert.AreEqual("Original", selected.Last(), "restore scenario " + scenario);
            }
            foreach (var list in new[] {
                new VbeDebugWindows.OptionsControl { Choices = new string[0], Value = "Original" },
                new VbeDebugWindows.OptionsControl { Choices = new[] { "Original", "Original" }, Value = "Original" },
                new VbeDebugWindows.OptionsControl { Choices = new[] { "Original" }, Value = "Unknown" },
                new VbeDebugWindows.OptionsControl { Choices = new[] { "Original" }, Value = "Original", Error = "unreadable" } })
            {
                int selects = 0;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CaptureOptionsFormatCategories(list, _ => selects++, () => null));
                Assert.AreEqual(0, selects);
            }
        }

        /// <summary>Preserves both inspection and restoration errors without claiming restored selection.</summary>
        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void FormatCategoryInspectionPreservesPrimaryAndRestorationFailures(bool selectionFails)
        {
            var list = new VbeDebugWindows.OptionsControl { Choices = new[] { "Original", "Other" }, Value = "Original" };
            var primary = new InvalidOperationException(selectionFails ? "Category selection failed" : "Palette read failed");
            var restoration = new System.IO.IOException("Original category could not be restored");
            var selections = new List<string>();
            int originalAttempts = 0;
            var palettes = new[] { "Foreground", "Background", "Indicator" }.Select(name =>
                new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.ComboBox", Value = "Automatic" }).ToArray();

            var error = Assert.ThrowsException<AggregateException>(() => VbeDebugWindows.CaptureOptionsFormatCategories(list, name => {
                selections.Add(name);
                if (name == "Original" && ++originalAttempts == 2) throw restoration;
                if (selectionFails && name == "Other") throw primary;
            }, () => { if (!selectionFails) throw primary; return palettes; }));

            Assert.AreEqual(2, error.InnerExceptions.Count);
            Assert.AreSame(primary, error.InnerExceptions[0]);
            Assert.AreSame(restoration, error.InnerExceptions[1]);
            CollectionAssert.AreEqual(selectionFails ? new[] { "Original", "Other", "Original" } : new[] { "Original", "Original" }, selections.ToArray());
        }

        /// <summary>A lone restoration failure remains the original exception after a successful inspection.</summary>
        [TestMethod]
        public void FormatCategoryInspectionPropagatesSoleRestorationFailure()
        {
            var list = new VbeDebugWindows.OptionsControl { Choices = new[] { "Original", "Other" }, Value = "Original" };
            var restoration = new System.IO.IOException("Restoration failed after complete inspection");
            var selections = new List<string>();
            var palettes = new[] { "Foreground", "Background", "Indicator" }.Select(name =>
                new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.ComboBox", Value = "Automatic" }).ToArray();
            var error = Assert.ThrowsException<System.IO.IOException>(() => VbeDebugWindows.CaptureOptionsFormatCategories(list,
                name => { selections.Add(name); if (selections.Count == 3) throw restoration; }, () => palettes));
            Assert.AreSame(restoration, error);
            CollectionAssert.AreEqual(new[] { "Original", "Other", "Original" }, selections.ToArray());
        }

        /// <summary>La version couvre les palettes d'autres catégories avant toute sélection ni écriture.</summary>
        [TestMethod]
        public void FormatCategoryColoursAreGloballyVersionedAndSelectedOnlyAfterGuard()
        {
            var probe = new FormatOptionsMatrixProbe(); var request = probe.Request();
            probe.Categories[1].Palettes[0].Value = "NativeIndex:2";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
            Assert.AreEqual(0, probe.Selections); Assert.AreEqual(0, probe.Inner.Writes);
            probe = new FormatOptionsMatrixProbe(); request = probe.Request();
            dynamic result = VbeDebugWindows.SetVbeOption(request, probe);
            Assert.AreEqual("Comment", result.Category); Assert.AreEqual("NativeIndex:1", result.After);
            Assert.AreEqual(1, probe.Selections); Assert.AreEqual(1, probe.Inner.Writes); Assert.AreEqual(1, probe.Inner.Accepts);
        }

        /// <summary>Catégorie absente, requête hors palette ou changement pendant l'écriture échouent avec Cancel.</summary>
        [TestMethod]
        public void FormatCategoryQueryGuardsWrongScopeUnknownCategoryAndChangedReadback()
        {
            foreach (int scenario in Enumerable.Range(0, 3))
            {
                var probe = new FormatOptionsMatrixProbe(); var request = probe.Request();
                if (scenario == 0) request.Query = "Unknown";
                if (scenario == 1) request.Property = "Code Colors";
                if (scenario == 2) probe.Inner.OnWrite = () => probe.Inner.Items[0].Value = "Normal";
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(scenario == 2 ? 1 : 0, probe.Inner.Writes);
                Assert.AreEqual(0, probe.Inner.Accepts); Assert.AreEqual(1, probe.Inner.Closes);
            }
            var noCategories = new WritableOptionsMatrixProbe(); var plain = noCategories.Request(); plain.Query = "Comment";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(plain, noCategories));
            Assert.AreEqual(0, noCategories.Writes);
        }

        /// <summary>Les palettes opaques exposent leurs vrais index, sans couleurs inventées, et versionnent les sélections.</summary>
        [TestMethod]
        public void NativeFormatChoicesPreserveUnlabeledIndicesAndInvalidateChangedSelections()
        {
            var control = new VbeDebugWindows.OptionsControl { Name = "Foreground", Type = "ControlType.ComboBox" };
            var palette = new[] { " Automatique" }.Concat(Enumerable.Repeat("", 16)).ToArray();
            VbeDebugWindows.DescribeOptionsNativeChoices(control, palette, 3, "");
            Assert.AreEqual(17, control.NativeChoices.Count); Assert.AreEqual(3, control.SelectedIndex);
            Assert.AreEqual("NativeIndex:3", control.Value); Assert.AreEqual("", control.NativeChoices[3].Label);
            Assert.AreEqual("NativeIndex:16", control.Choices[16]);
            Assert.AreEqual("NativeIndex:3", VbeDebugWindows.ValidateEditableOption("Editor Format", control, "NativeIndex:3"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ValidateEditableOption("Editor Format", control, "RGB:#ff0000"));
            foreach (bool changeLabel in new[] { false, true })
            {
                var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "Editor Format"; probe.Items[0] = control;
                var request = probe.Request(); request.Value = "NativeIndex:3";
                if (changeLabel) control.NativeChoices[3].Label = "changed native label";
                else control.SelectedIndex = 4;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(0, probe.Writes); Assert.AreEqual(1, probe.Closes);
                VbeDebugWindows.DescribeOptionsNativeChoices(control, palette, 3, "");
            }
        }

        /// <summary>Police non sélectionnée reste lisible; taille sans catalogue et libellés doubles ne sont pas inscriptibles.</summary>
        [TestMethod]
        public void NativeFormatTextReadsExactEditValueButRequiresUniqueEnumeratedChoiceForWriting()
        {
            var font = new VbeDebugWindows.OptionsControl { Name = "Font", Type = "ControlType.ComboBox" };
            VbeDebugWindows.DescribeOptionsNativeChoices(font, new[] { "Courier New", "Arial" }, -1, "Courier New");
            Assert.AreEqual(-1, font.SelectedIndex); Assert.AreEqual("Courier New", font.Value);
            Assert.AreEqual("Arial", VbeDebugWindows.ValidateEditableOption("Editor Format", font, "Arial"));
            VbeDebugWindows.DescribeOptionsNativeChoices(font, new[] { "Arial", "Arial" }, -1, "Courier New");
            Assert.AreEqual("Courier New", font.Value);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ValidateEditableOption("Editor Format", font, "Arial"));
            var size = new VbeDebugWindows.OptionsControl { Name = "Size", Type = "ControlType.ComboBox" };
            VbeDebugWindows.DescribeOptionsNativeChoices(size, new string[0], -1, "10");
            Assert.AreEqual("10", size.Value); Assert.AreEqual(0, size.Choices.Count);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ValidateEditableOption("Editor Format", size, "12"));
            foreach (int index in new[] { -2, 2 })
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, new[] { "A", "B" }, index, "A"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, null, -1, "A"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, new string[2001], -1, "A"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, new[] { new string('x', 4097) }, -1, "A"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, new string[] { null }, -1, "A"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.DescribeOptionsNativeChoices(font, new string[0], -1, null));
        }

        [TestMethod]
        public void DocumentedOptionsCompleteTheWholeVersionWriteReadbackCommitWorkflow()
        {
            foreach (var row in OptionsCompletionMatrix.Supported())
            {
                var probe = new WritableOptionsMatrixProbe();
                probe.Names[0] = row.Item1;
                var control = probe.Items[0];
                control.Name = row.Item2; control.Type = row.Item3;
                control.Value = row.Item3 == "ControlType.CheckBox" ? "On" : "Original";
                control.Choices = new[] { "Original", "Native choice" };
                var request = probe.Request(); request.Value = row.Item4;
                dynamic result = VbeDebugWindows.SetVbeOption(request, probe);
                Assert.IsTrue((bool)result.ControlValueVerified, row.Item2);
                Assert.IsTrue((bool)result.DialogClosed, row.Item2);
                Assert.IsFalse((bool)result.PersistenceVerified, row.Item2);
                Assert.AreEqual(1, probe.Writes); Assert.AreEqual(1, probe.Accepts);
                Assert.AreEqual(0, probe.Closes);
            }
        }

        [TestMethod]
        public void GridDimensionsRejectInvalidValuesAndCancelBeforeWriting()
        {
            foreach (string name in new[] { "Width", "Height", "Largeur", "Hauteur" })
                foreach (object value in OptionsCompletionMatrix.InvalidGridValues)
                {
                    var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "General";
                    probe.Items[0].Name = name; probe.Items[0].Type = "ControlType.Edit"; probe.Items[0].Value = "6";
                    var request = probe.Request(); request.Value = value;
                    Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                    Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
                }
        }

        [TestMethod]
        public void FormatChoicesRejectUnknownAmbiguousUnreadableOrArbitraryValues()
        {
            foreach (string scenario in OptionsCompletionMatrix.InvalidChoiceScenarios)
            {
                var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "Editor Format";
                var control = probe.Items[0]; control.Name = "Font"; control.Type = "ControlType.ComboBox";
                control.Value = "Consolas"; control.Choices = new[] { "Consolas", "Courier New" };
                object value = "Courier New";
                if (scenario == "absent") value = "invented font";
                if (scenario == "duplicate") control.Choices = new[] { "Courier New", "Courier New" };
                if (scenario == "null catalogue") control.Choices = null;
                if (scenario == "empty catalogue") control.Choices = new string[0];
                if (scenario == "null value") value = null;
                if (scenario == "numeric value") value = 12;
                if (scenario == "wrong tab") probe.Names[0] = "General";
                if (scenario == "unknown name") control.Name = "Unknown";
                if (scenario == "wrong type") control.Type = "ControlType.Edit";
                if (scenario == "unreadable") control.Error = "Selection unavailable";
                var request = probe.Request(); request.Value = value;
                if (scenario == "null value" || scenario == "numeric value")
                    Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                else Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
            }
        }

        [TestMethod]
        public void ChoiceCatalogueSelectionAndColorCategoryAreAllVersioned()
        {
            foreach (string scenario in OptionsCompletionMatrix.RevisionScenarios)
            {
                var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "Editor Format";
                var color = probe.Items[0]; color.Name = "Foreground"; color.Type = "ControlType.ComboBox";
                color.Value = "Black"; color.Choices = new[] { "Black", "Red" };
                var category = new VbeDebugWindows.OptionsControl { Name = "Code Colors", Type = "ControlType.List",
                    Value = "Normal Text", Choices = new[] { "Normal Text", "Comment Text" } };
                probe.Items.Add(category);
                var request = probe.Request(); request.Value = "Red";
                if (scenario == "choices") color.Choices = new[] { "Black", "Blue" };
                if (scenario == "selection") color.Value = "Red";
                if (scenario == "category") category.Value = "Comment Text";
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
            }
        }

        [TestMethod]
        public void FormatChoiceReadbackFailureCancelsWithoutAccepting()
        {
            var probe = new WritableOptionsMatrixProbe { IgnoreWrite = true }; probe.Names[0] = "Editor Format";
            probe.Items[0].Name = "Size"; probe.Items[0].Type = "ControlType.ComboBox";
            probe.Items[0].Value = "10"; probe.Items[0].Choices = new[] { "10", "12" };
            var request = probe.Request(); request.Value = "12";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
            Assert.AreEqual(1, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        public void SameNamedLabelsAreIgnoredAndOnlyEditableControlsCanBeWritten()
        {
            foreach (string kind in new[] { "ControlType.Text", "ControlType.Button", null })
            {
                var probe = new WritableOptionsMatrixProbe();
                var option = probe.Items[0];
                probe.Items.Add(new VbeDebugWindows.OptionsControl { Name = option.Name, Type = kind,
                    Visible = true, Enabled = true, Value = option.Value });
                var request = probe.Request();
                Assert.IsTrue((bool)((dynamic)VbeDebugWindows.SetVbeOption(request, probe)).ControlValueVerified);
                Assert.AreEqual(1, probe.Writes);

                var labelOnly = new WritableOptionsMatrixProbe();
                labelOnly.Items[0].Type = kind;
                var absent = labelOnly.Request();
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(absent, labelOnly));
                Assert.AreEqual(0, labelOnly.Writes);
                Assert.AreEqual(1, labelOnly.Closes);
            }
        }

        [TestMethod]
        public void WritableOptionsGuardEveryRequiredFieldTabCatalogueAndExactVisibleControl()
        {
            foreach(int scenario in Enumerable.Range(0,5))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)r=null;if(scenario==1)r.Pane=" ";if(scenario==2)r.Property=null;if(scenario==3)r.ExpectedOptionsVersion="";if(scenario==4)p.Open=false;
                if(scenario<4)Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.SetVbeOption(r,p));
                else Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));
                Assert.AreEqual(0,p.Writes);
            }
            foreach(int scenario in Enumerable.Range(0,9))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)p.Names.Clear();if(scenario==1)p.Names.AddRange(Enumerable.Repeat("Editor",8));if(scenario==2)p.Names[0]=" ";
                if(scenario==3)p.Items.AddRange(Enumerable.Repeat(p.Items[0],2000));
                if(scenario==4)r.Pane="Missing";
                if(scenario==5){int calls=0;p.OnTabs=()=>{if(++calls==2)p.Names.Add("Editor");};}
                if(scenario==6)r.Property="Missing";
                if(scenario==7){p.Items[0].Error="unreadable";r=p.Request();}
                if(scenario==8){p.Items.Add(p.Items[0]);r=p.Request();}
                Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));Assert.AreEqual(0,p.Writes);Assert.AreEqual(1,p.Closes);
            }
        }
        [TestMethod]
        public void WritableOptionsKnownReadbackFailureCancelsButUncertainAcceptRetainsDialog()
        {
            foreach(int scenario in Enumerable.Range(0,7))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)p.OnWrite=()=>p.Items.Clear();
                if(scenario==1)p.OnWrite=()=>p.Items.Add(new VbeDebugWindows.OptionsControl{Name=r.Property,Type="ControlType.CheckBox",Value="Off"});
                if(scenario==2)p.OnWrite=()=>p.Items[0].Error="readback failed";
                if(scenario==3)p.OnWrite=()=>p.Items[0].Type="other";
                if(scenario==4)p.OnWrite=()=>p.Items[0].Visible=false;
                if(scenario==5)p.OnWrite=()=>p.Items[0].Enabled=false;
                if(scenario==6)p.OnAccept=()=>{throw new InvalidOperationException("OK rejected");};
                p.IgnoreWrite=scenario<6;
                Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));
                Assert.AreEqual(scenario == 6 ? 0 : 1,p.Closes);
                Assert.AreEqual(scenario == 6,p.Open);
            }
            var open=new WritableOptionsMatrixProbe{KeepOpen=true};var request=open.Request();
            dynamic result=VbeDebugWindows.SetVbeOption(request,open);Assert.IsFalse((bool)result.DialogClosed);Assert.AreEqual(0,open.Closes);Assert.IsTrue(open.Open);
            foreach(string type in new[]{"ControlType.RadioButton","ControlType.Edit"})
            {
                var p=new WritableOptionsMatrixProbe();p.Names[0]=type.EndsWith("Edit")?"Editor":"General";
                p.Items[0].Type=type;p.Items[0].Name=type.EndsWith("Edit")?"Tab Width":"Break on All Errors";p.Items[0].Value=type.EndsWith("Edit")?(object)"4":false;
                var r=p.Request();r.Value=type.EndsWith("Edit")?(object)8:true;
                Assert.IsTrue((bool)((dynamic)VbeDebugWindows.SetVbeOption(r,p)).ControlValueVerified);
            }
        }
        [TestMethod]
        public void PreferenceAllowlistCoversLocalizedTabsTypesAndNullableOrInvalidValues()
        {
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.SetVbeOption(null));
            var p=new WritableOptionsMatrixProbe();
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Hidden",Type="ControlType.CheckBox",Visible=false,Enabled=true});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Disabled",Type="ControlType.CheckBox",Visible=true,Enabled=false});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Other",Type="ControlType.CheckBox",Visible=true,Enabled=true});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name=" ",Type="ControlType.Text",Visible=true,Enabled=true});
            var on=p.Request();on.Value=true;
            Assert.AreEqual("On",(string)((dynamic)VbeDebugWindows.SetVbeOption(on,p)).After);
            var check=new VbeDebugWindows.OptionsControl{Type="ControlType.CheckBox",Name="&Auto Syntax Check:"};
            foreach(string tab in new[]{"Editor","éditeur","editeur"})Assert.AreEqual(true,VbeDebugWindows.ValidateEditableOption(tab,check,true));
            check.Name="Compile on demand";foreach(string tab in new[]{"General","général"})Assert.AreEqual(false,VbeDebugWindows.ValidateEditableOption(tab,check,false));
            var radio=new VbeDebugWindows.OptionsControl{Type="ControlType.RadioButton",Name="Break on Unhandled Errors"};
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("General",radio,"true"));
            foreach(string tab in new[]{null,"Editor Format"})Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption(tab,new VbeDebugWindows.OptionsControl(),true));
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",radio,true));
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("General",new VbeDebugWindows.OptionsControl{Type="ControlType.Edit",Name="Tab Width"},4));
        }
        [TestMethod]
        public void OptionsRevisionProtectsMutationAndRequestedValueIsReadBeforeCommit()
        {
            var fake = new Fake();
            dynamic read = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            var request = new Request { Pane = "Editor", Property = "Auto Syntax Check", Value = false, ExpectedOptionsVersion = (string)read.OptionsVersion };
            dynamic result = VbeDebugWindows.SetVbeOption(request, fake);
            Assert.IsTrue((bool)result.ControlValueVerified); Assert.IsTrue((bool)result.DialogClosed);
            Assert.IsFalse((bool)result.PersistenceVerified); Assert.AreEqual(1, fake.Writes); Assert.AreEqual(1, fake.Accepts);
            Assert.AreEqual("On", (string)result.Before); Assert.AreEqual("Off", (string)result.After);
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, fake));
            Assert.AreEqual(1, fake.Writes); Assert.AreEqual(1, fake.Accepts); Assert.IsFalse(fake.Open);
        }
        [TestMethod]
        public void OptionsRefuseUnsupportedPreferencesAndCancelFailedReadback()
        {
            var fake = new Fake(); dynamic read = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            var request = new Request { Pane="Editor",Property="Auto Syntax Check",Value=false,ExpectedOptionsVersion=(string)read.OptionsVersion };
            fake.RetainWrite = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, fake));
            Assert.AreEqual(0, fake.Accepts); Assert.IsFalse(fake.Open);
            var check = new VbeDebugWindows.OptionsControl { Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On" };
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor Format",check,false));
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",check,"false"));
            check.Name="Unknown preference";
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",check,true));
            var width = new VbeDebugWindows.OptionsControl { Name="Tab Width",Type="ControlType.Edit",Value="4" };
            Assert.AreEqual("8",VbeDebugWindows.ValidateEditableOption("Editor",width,8));
            foreach(object invalid in new object[] {0,33,1.5,"4: Stop",true})
                Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",width,invalid));
            var radio = new VbeDebugWindows.OptionsControl { Name="Break on All Errors",Type="ControlType.RadioButton",Value=false };
            Assert.AreEqual(true,VbeDebugWindows.ValidateEditableOption("General",radio,true));
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("General",radio,false));
        }
        [TestMethod]
        public void FrenchNativeLabelsAndTabWidthLabelDoNotPreventEdits()
        {
            foreach (string name in new[] { "Complément automatique des instructions", "Info express automatique" })
                Assert.AreEqual(false, VbeDebugWindows.ValidateEditableOption("Éditeur",
                    new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.CheckBox" }, false));
            Assert.AreEqual(false, VbeDebugWindows.ValidateEditableOption("Général",
                new VbeDebugWindows.OptionsControl { Name = "Compilation sur demande", Type = "ControlType.CheckBox" }, false));
            var fake = new Fake { Width = true };
            dynamic before = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            dynamic after = VbeDebugWindows.SetVbeOption(new Request { Pane = "Editor", Property = "Largeur de la tabulation :",
                Value = 32, ExpectedOptionsVersion = (string)before.OptionsVersion }, fake);
            Assert.AreEqual("32", (string)after.After); Assert.AreEqual(1, fake.Accepts);
        }
        [TestMethod]
        public void DuplicateDebugObservationsKeepDistinctPathsAndValues()
        {
            string first = VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "values", "child" });
            Assert.AreEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "values", "child" }));
            Assert.AreNotEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 6 Type Long", new[] { "values", "child" }));
            Assert.AreNotEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "other", "child" }));
            Assert.AreNotEqual(VbeDebugWindows.DebugRowIdentity("row", new[] { "a/b", "c" }), VbeDebugWindows.DebugRowIdentity("row", new[] { "a", "b/c" }));
        }
        private sealed class Fake : VbeDebugWindows.IWritableOptionsProbe
        {
            public bool Open = true, RetainWrite = true, Width;
            public int Writes, Accepts;
            private readonly VbeDebugWindows.OptionsControl choice = new VbeDebugWindows.OptionsControl { Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On" };
            public IntPtr Dialog() => Open ? new IntPtr(1) : IntPtr.Zero;
            public IList<string> Tabs(IntPtr dialog) => new[] { "Editor" };
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog,int index) {
                if (!Width) return new[] { choice };
                choice.Name = "Largeur de la tabulation :"; choice.Type = "ControlType.Edit";
                if (choice.Value is string text && (text == "On" || text == "Off")) choice.Value = "4";
                return new[] { new VbeDebugWindows.OptionsControl { Name = choice.Name, Type = "ControlType.Text" }, choice };
            }
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Close(IntPtr dialog) { Open=false; }
            public void Pause(int milliseconds) { }
            public void Write(IntPtr dialog,int index,string name,string type,object value) { Writes++; if(RetainWrite)choice.Value=Width ? value : (object)((bool)value ? "On":"Off"); }
            public void Accept(IntPtr dialog) { Accepts++;Open=false; }
        }
    }
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeOptionReadAndWriteCoverEmbeddedEditsAndNonNativeComboProviders()
        {
            using (var f = new OwnedNativeOptionsControls())
            {
                IntPtr edit = IntPtr.Zero;
                f.Host.Invoke(owner => edit = OptionsFixtureCreate(0, "Edit", "", 0x50000000, 0, 0, 30, 15, f.Font, new IntPtr(519), IntPtr.Zero, IntPtr.Zero));
                Assert.AreNotEqual(IntPtr.Zero, edit);
                f.Root.Add(new AutomationNode { Name = "Embedded font edit", Kind = ControlType.Edit, NativeHandle = edit.ToInt32() }.With(ValuePattern.Pattern));
                IntPtr standalone = IntPtr.Zero;
                f.Host.Invoke(owner => standalone = OptionsFixtureCreate(0, "Edit", "text", 0x50000000, 0, 0, 30, 15, f.Host.Handle, new IntPtr(520), IntPtr.Zero, IntPtr.Zero));
                Assert.AreNotEqual(IntPtr.Zero, standalone);
                f.Root.Add(new AutomationNode { Name = "Standalone native edit", Kind = ControlType.Edit, NativeHandle = standalone.ToInt32(), Text = "text" }.With(ValuePattern.Pattern));
                IntPtr nonCombo = IntPtr.Zero;
                f.Host.Invoke(owner => nonCombo = OptionsFixtureCreate(0, "Static", "Provider", 0x50000000, 0, 0, 30, 15, f.Host.Handle, new IntPtr(521), IntPtr.Zero, IntPtr.Zero));
                Assert.AreNotEqual(IntPtr.Zero, nonCombo);
                foreach (int handle in new[] { 0, nonCombo.ToInt32() })
                {
                    var combo = f.Root.Add(new AutomationNode { Name = "Provider " + handle, Kind = ControlType.ComboBox, NativeHandle = handle }.With(SelectionPattern.Pattern));
                    var choice = combo.Add(new AutomationNode { Name = "Choice", Selected = true }.With(SelectionItemPattern.Pattern));
                    var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(f.Host.Handle);
                    var controls = probe.Controls(f.Host.Handle, 0);
                    Assert.IsFalse(controls.Any(c => c.Name == "Embedded font edit"));
                    Assert.IsTrue(controls.Any(c => c.Name == "Standalone native edit" && Equals(c.Value, "text")));
                    ((VbeDebugWindows.IWritableOptionsProbe)probe).Write(f.Host.Handle, 0, combo.Name, "ControlType.ComboBox", "Choice");
                    Assert.IsTrue(choice.Selected);
                }
                var racing = f.Root.Add(new AutomationNode { Name = "Password race", Kind = ControlType.ComboBox, Text = "must not read" }.With(ValuePattern.Pattern));
                int reads = 0; racing.OnPropertyRead = id => { if (id == AutomationElement.IsPasswordProperty.Id && ++reads >= 2) racing.Password = true; };
                var reader = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); reader.Tabs(f.Host.Handle);
                var refused = reader.Controls(f.Host.Handle, 0).Single(c => c.Name == racing.Name);
                Assert.IsNull(refused.Value); Assert.IsNotNull(refused.Error);
            }
            Assert.AreEqual("", InvokeOptionsMethod(null, "NormalizeOptionName", new object[] { null }));
            var nativeType = typeof(VbeDebugWindows).GetNestedType("NativeProbe", System.Reflection.BindingFlags.NonPublic);
            var ctor = nativeType.GetConstructor(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, new[] { typeof(int) }, null);
            var failure = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => ctor.Invoke(new object[] { 0 }));
            Assert.IsInstanceOfType(failure.InnerException, typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void NativeCategoryLateIdentityPatternAndReadbackChangesAreRefused()
        {
            foreach (string scenario in new[] { "pid", "pattern", "identifier", "uia-handle", "selection", "palette-selection" })
            using (var f = new OwnedNativeOptionsControls())
            {
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(f.Host.Handle);
                var categories = (VbeDebugWindows.IFormatCategoriesOptionsProbe)probe;
                var target = f.CategoryItems[1]; int reads = 0;
                if (scenario == "pid") target.OnPropertyRead = id => { if (id == AutomationElement.ProcessIdProperty.Id && ++reads >= 2) target.ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id + 1; };
                if (scenario == "pattern") target.OnPatternRead = id => { if (id == SelectionItemPattern.Pattern.Id && ++reads >= 2) target.Patterns.Remove(id); };
                var select = target.SelectedAction;
                if (scenario == "identifier") target.SelectedAction = () => { select(); f.Host.Invoke(owner => OptionsFixtureSetStyle(f.List, -12, 999)); };
                if (scenario == "uia-handle") target.SelectedAction = () => { select(); f.Categories.NativeHandle = f.Font.ToInt32(); };
                if (scenario == "selection") target.SelectedAction = () => { target.Selected = false; f.CategoryItems[0].Selected = true; };
                if (scenario == "palette-selection") f.Palettes[0].OnPropertyRead = id => { if (id == AutomationElement.NameProperty.Id && f.CurrentCategory == "Comment") { target.Selected = false; f.CategoryItems[0].Selected = true; } };
                if (scenario == "palette-selection") Assert.ThrowsException<InvalidOperationException>(() => categories.FormatCategories(f.Host.Handle, 0), scenario);
                else Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(f.Host.Handle, 0, "Comment"), scenario);
            }
        }
        [TestMethod]
        public void NativeComboWriteUsesItsExactHandleAndRefusesLateParentLoss()
        {
            using (var f = new OwnedNativeOptionsControls())
            {
                f.Host.Invoke(owner => f.Root.Add(new AutomationNode { Name = "Font", Kind = ControlType.ComboBox, NativeHandle = f.Font.ToInt32() }));
                var probe = Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");
                probe.Tabs(f.Host.Handle);
                probe.Write(f.Host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New");
                Assert.AreEqual(1, OptionsFixtureInteger(f.Font, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32());
                ParentLostOnComboSelection lost = null;
                f.Host.Invoke(owner => { lost = new ParentLostOnComboSelection(f.Font) { Armed = true }; });
                try
                {
                    var error = Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "WriteOptionsCombo", f.Font, "Consolas"));
                    StringAssert.Contains(error.Message, "parent does not belong to this process");
                    Assert.AreEqual(f.Host.Handle, lost.PreviousParent);
                    Assert.AreEqual(IntPtr.Zero, InvokeOptionsMethod(null, "OptionsComboParent", f.Font));
                }
                finally { f.Host.Invoke(owner => lost.Dispose()); }
            }
        }
        [TestMethod]
        public void NativeOptionsEnumerateAndSelectExactChoicesWithoutTypingOrInventingColors()
        {
            var root = new AutomationNode { Name = "Options", Kind = System.Windows.Automation.ControlType.Window };
            root.Add(new AutomationNode { Name = "Editor Format", Kind = System.Windows.Automation.ControlType.TabItem }
                .With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var combo = root.Add(new AutomationNode { Name = "Font", Kind = System.Windows.Automation.ControlType.ComboBox }
                .With(System.Windows.Automation.SelectionPattern.Pattern, System.Windows.Automation.ValuePattern.Pattern));
            var original = combo.Add(new AutomationNode { Name = "Consolas", Selected = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var desired = combo.Add(new AutomationNode { Name = "Courier New", Offscreen = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            desired.SelectedAction = () => original.Selected = false;
            using (var host = new AutomationHost(root, optionsDialog: true))
            using (var scene = new SystemScene())
            {
                BindOwnedOptionsDialog(scene, host);
                var native = Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe"); native.Tabs(host.Handle);
                var observed = native.Controls(host.Handle, 0).Single(x => x.Name == "Font");
                Assert.AreEqual("Consolas", observed.Value);
                CollectionAssert.AreEqual(new[] { "Consolas", "Courier New" }, observed.Choices.ToArray());
                native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New");
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
                Assert.AreEqual("", combo.Text); Assert.AreEqual(0, combo.FocusCount);
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "invented"));
                desired.Enabled = false;
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New")); desired.Enabled = true;
                desired.Patterns.Clear();
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                desired.Patterns.Add(System.Windows.Automation.SelectionItemPattern.Pattern.Id);
                var duplicate = combo.Add(new AutomationNode { Name = "Courier New" }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                combo.Children.Remove(duplicate);
                combo.Password = true;
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); combo.Password = false;
                original.Selected = true;
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); original.Selected = false;
                desired.Selected = false;
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); desired.Selected = true;
                combo.Patterns.Remove(System.Windows.Automation.SelectionPattern.Pattern.Id); combo.Text = "Courier New";
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
                combo.Patterns.Clear();
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error));
                combo.Kind = System.Windows.Automation.ControlType.List;
                combo.Patterns.Add(System.Windows.Automation.SelectionPattern.Pattern.Id);
                native.Write(host.Handle, 0, "Font", "ControlType.List", "Courier New");
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
            }
        }

        [TestMethod]
        public void NativeOptionsRefusesHandlelessCheckboxAndPreservesOtherUiaPatterns()
        {
            var root=new AutomationNode{Name="Options",Kind=System.Windows.Automation.ControlType.Window};
            root.Add(new AutomationNode{Name="Editor",Kind=System.Windows.Automation.ControlType.TabItem}.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var check=root.Add(new AutomationNode{Name="Auto Syntax Check",Kind=System.Windows.Automation.ControlType.CheckBox}.With(System.Windows.Automation.TogglePattern.Pattern));
            var radio=root.Add(new AutomationNode{Name="Break on All Errors",Kind=System.Windows.Automation.ControlType.RadioButton}.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var edit=root.Add(new AutomationNode{Name="Tab Width",Kind=System.Windows.Automation.ControlType.Edit,Text="4"}.With(System.Windows.Automation.ValuePattern.Pattern));
            using(var host=new AutomationHost(root, optionsDialog:true))
            using(var scene=new SystemScene())
            {
                BindOwnedOptionsDialog(scene, host);
                var native=Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");native.Tabs(host.Handle);
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));
                Assert.AreEqual(System.Windows.Automation.ToggleState.Off,check.ToggleState, "A handleless checkbox must not fall back to UIA Toggle.");
                native.Write(host.Handle,0,radio.Name,"ControlType.RadioButton",true);Assert.IsTrue(radio.Selected);
                native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8);Assert.AreEqual("8",edit.Text);
                check.ToggleState=System.Windows.Automation.ToggleState.Indeterminate;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.ToggleState=System.Windows.Automation.ToggleState.Off;
                edit.ReadOnly=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8));edit.ReadOnly=false;
                foreach(var node in new[]{check,radio,edit})
                {
                    var pattern=node==check?System.Windows.Automation.TogglePattern.Pattern:node==radio?System.Windows.Automation.SelectionItemPattern.Pattern:System.Windows.Automation.ValuePattern.Pattern;
                    node.Patterns.Clear();Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,node.Name,node.Kind.ProgrammaticName,true));node.Patterns.Add(pattern.Id);
                }
                edit.Password=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8));edit.Password=false;
                Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,"absent","ControlType.Edit",8));
                check.Offscreen=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.Offscreen=false;
                check.Enabled=false;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.Enabled=true;
                Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.Edit",8));
                root.Add(new AutomationNode{Name=check.Name,Kind=System.Windows.Automation.ControlType.CheckBox});Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));root.Children.RemoveAt(root.Children.Count-1);
                var slider=root.Add(new AutomationNode{Name="Slider",Kind=System.Windows.Automation.ControlType.Slider});Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,slider.Name,"ControlType.Slider",8));
            }
        }
        [TestMethod]
        public void NativeCheckboxUsesOwnedWin32ButtonAndNeverTheUiaToggleProvider()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                IntPtr button = IntPtr.Zero;
                fixture.Host.Invoke(owner => button = OptionsFixtureCreate(0, "Button", "", 0x50000003,
                    10, 10, 160, 25, owner, new IntPtr(540), IntPtr.Zero, IntPtr.Zero));
                Assert.AreNotEqual(IntPtr.Zero, button);
                var node = fixture.Root.Add(new AutomationNode { Name = "Margin Indicator Bar", Kind = ControlType.CheckBox,
                    NativeHandle = button.ToInt32(), ToggleState = ToggleState.Off }.With(TogglePattern.Pattern));
                var probe = Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");
                probe.Tabs(fixture.Host.Handle);
                probe.Write(fixture.Host.Handle, 0, node.Name, "ControlType.CheckBox", true);
                Assert.AreEqual(1, OptionsFixtureInteger(button, 0xF0, IntPtr.Zero, IntPtr.Zero).ToInt32());
                Assert.AreEqual(ToggleState.Off, node.ToggleState, "The fake provider must never receive Toggle.");
                probe.Write(fixture.Host.Handle, 0, node.Name, "ControlType.CheckBox", true);
                probe.Write(fixture.Host.Handle, 0, node.Name, "ControlType.CheckBox", false);
                Assert.AreEqual(0, OptionsFixtureInteger(button, 0xF0, IntPtr.Zero, IntPtr.Zero).ToInt32());
                Assert.AreEqual(2, fixture.Notifications.Count(x => x.Item1 == 540), "Only the two real transitions notify the parent.");
            }
        }

        [TestMethod]
        public void NativeOptionsAcceptChecksEveryOkButtonBoundaryBeforeSendingOwnedClick()
        {
            var saved=VbeDebugWindows.OptionsWindowEnabled;
            try
            {
                using(var scene=new SystemScene())
                {
                    var native=Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");var dialog=scene.Add("Options");
                    Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    var ok=scene.Add("OK","Edit",dialog,1);Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    ok.Class="Button";VbeDebugWindows.OptionsWindowEnabled=handle=>false;Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    VbeDebugWindows.OptionsWindowEnabled=handle=>true;scene.PostSucceeds=false;Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    scene.PostSucceeds=true;int clicks=0;scene.OnMessage=(window,message)=>{if(window==ok&&message==0xF5)clicks++;};native.Accept(dialog.Handle);Assert.AreEqual(1,clicks);
                }
            }
            finally{VbeDebugWindows.OptionsWindowEnabled=saved;}
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Windows.Automation;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        /// <summary>Les vraies listes ComboBox détenues livrent leurs valeurs/indices, écrivent une entrée exacte et notifient le propriétaire.</summary>
        [TestMethod]
        public void OwnedNativeOptionsCombosReadAndSelectExactValuesWithoutTyping()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var font = new VbeDebugWindows.OptionsControl();
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Font, font);
                Assert.AreEqual("Consolas", font.Value); Assert.AreEqual(-1, font.SelectedIndex);
                CollectionAssert.AreEqual(new[] { "Consolas", "Courier New", "Duplicate", "Duplicate" }, font.Choices.ToArray());
                InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Font, "Courier New");
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Font, font);
                Assert.AreEqual("Courier New", font.Value); Assert.AreEqual(1, font.SelectedIndex);
                Assert.IsTrue(fixture.Notifications.Any(x => x.Item1 == 510 && x.Item2 == 1 && x.Item3 == fixture.Font));
                Assert.IsTrue(fixture.Notifications.Any(x => x.Item1 == 510 && x.Item2 == 9 && x.Item3 == fixture.Font));
                foreach (string choice in new[] { "Missing", "Duplicate" })
                    Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Font, choice));
                var size = new VbeDebugWindows.OptionsControl();
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, size);
                Assert.AreEqual("10", size.Value); Assert.AreEqual(0, size.Choices.Count);
                Assert.AreEqual(IntPtr.Zero, OptionsFixtureInteger(fixture.Size, 0x157, IntPtr.Zero, IntPtr.Zero), "Read must close its temporary dropdown.");
                var palette = new VbeDebugWindows.OptionsControl(); var handle = new IntPtr(fixture.Palettes[0].NativeHandle.Value);
                InvokeOptionsMethod(null, "WriteOptionsCombo", handle, "NativeIndex:2");
                InvokeOptionsMethod(null, "ReadOptionsCombo", handle, palette);
                Assert.AreEqual("NativeIndex:2", palette.Value); Assert.AreEqual(2, palette.SelectedIndex);
                Assert.AreEqual("", palette.NativeChoices[2].Label); Assert.AreEqual(2, fixture.Colours["Normal"][0]);
            }
        }

        [TestMethod]
        public void OwnedNativeEmptySizeTraceRetainsCountsIdentityAndClosureWithoutValues()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var rows = new List<string>();
                var trace = new VbeInspectionTrace(rows.Add);
                var size = new VbeDebugWindows.OptionsControl { Name = "Taille :" };
                using (trace.Enter()) InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, size);
                Assert.AreEqual("10", size.Value); Assert.AreEqual(0, size.Choices.Count);
                Assert.AreEqual(1, rows.Count);
                var row = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(rows[0]);
                var native = (Dictionary<string, object>)row["Native"];
                Assert.AreEqual("NativeCombo", row["Reader"]); Assert.AreEqual("Size", row["Role"]);
                Assert.AreEqual(511, native["ControlId"]); Assert.AreEqual(0, native["CountBefore"]);
                Assert.AreEqual(0, native["CountAfterExpansion"]); Assert.AreEqual(true, native["ExpansionAttempted"]);
                Assert.AreEqual(false, native["DropDownBefore"]); Assert.AreEqual(false, native["DropDownAfterCleanup"]);
                Assert.AreEqual(true, native["ReadCompleted"]);
                Assert.AreEqual(fixture.Size.ToInt64(), Convert.ToInt64(native["Window"]));
                Assert.AreEqual(fixture.Host.Handle.ToInt64(), Convert.ToInt64(native["Parent"]));
                Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().Id, Convert.ToInt32(native["OwnerProcessId"]));
                Assert.IsTrue(Convert.ToInt32(native["OwnerThreadId"]) > 0);
                Assert.IsTrue((Convert.ToInt32(native["Style"]) & 0x200) != 0);
                Assert.IsFalse(native.ContainsKey("Value")); Assert.IsFalse(native.ContainsKey("Choices"));
                Assert.IsFalse(rows[0].Contains("Taille"));
                Assert.IsFalse(fixture.Notifications.Any(x => x.Item1 == 511 && (x.Item2 == 1 || x.Item2 == 9)));
            }
        }

        [TestMethod]
        public void OwnedNativeLazySizeTraceObservesPopulationWithoutASecondExpansion()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                int openings = 0;
                fixture.OnControlNotification = (identifier, code, window) => {
                    if (identifier != 511 || code != 7) return;
                    openings++;
                    foreach (string size in new[] { "8", "10", "12" }) OptionsFixtureText(window, 0x143, IntPtr.Zero, size);
                };
                var rows = new List<string>();
                var control = new VbeDebugWindows.OptionsControl { Name = "Size" };
                using (new VbeInspectionTrace(rows.Add).Enter()) InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, control);
                Assert.AreEqual(1, openings);
                CollectionAssert.AreEqual(new[] { "8", "10", "12" }, control.Choices.ToArray());
                Assert.AreEqual("10", control.Value);
                var row = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(rows.Single());
                var native = (Dictionary<string, object>)row["Native"];
                Assert.AreEqual(0, native["CountBefore"]); Assert.AreEqual(3, native["CountAfterExpansion"]);
                Assert.AreEqual(false, native["DropDownAfterCleanup"]);
                Assert.IsFalse(fixture.Notifications.Any(x => x.Item1 == 511 && (x.Item2 == 1 || x.Item2 == 9)));
            }
        }

        [TestMethod]
        public void OwnedNativeSizeCataloguePopulatesOnFocusWithoutSelectionOrEditWrite()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                int focusNotifications = 0;
                fixture.OnControlNotification = (identifier, code, window) => {
                    if (identifier != 511 || code != 3) return; // CBN_SETFOCUS, as observed in the native diagnostic.
                    focusNotifications++;
                    foreach (string size in new[] { "8", "10", "12" }) OptionsFixtureText(window, 0x143, IntPtr.Zero, size);
                };
                var rows = new List<string>();
                var control = new VbeDebugWindows.OptionsControl { Name = "Size" };
                using (new VbeInspectionTrace(rows.Add).Enter()) InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, control);
                CollectionAssert.AreEqual(new[] { "8", "10", "12" }, control.Choices.ToArray());
                Assert.AreEqual("10", control.Value); Assert.AreEqual(1, focusNotifications);
                var native = (Dictionary<string, object>)new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(rows.Single())["Native"];
                Assert.AreEqual(0, native["CountBefore"]); Assert.AreEqual(3, native["CountAfterFocus"]);
                Assert.AreEqual(true, native["FocusAttempted"]); Assert.AreEqual(false, native["ExpansionAttempted"]);
                Assert.IsFalse(fixture.Notifications.Any(x => x.Item1 == 511 && (x.Item2 == 1 || x.Item2 == 9)));
            }
        }

        [TestMethod]
        public void OwnedNativeSizeIdentifierPreparesCatalogueForUnlabelledWriteVerification()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                OptionsFixtureSetStyle(fixture.Size, -12, 4911); // The observed VBE Size identifier, not a guessed size value.
                fixture.OnControlNotification = (identifier, code, window) => {
                    if (identifier != 4911 || code != 3) return;
                    foreach (string size in new[] { "8", "10", "12" }) OptionsFixtureText(window, 0x143, IntPtr.Zero, size);
                };
                var control = new VbeDebugWindows.OptionsControl(); // Same verifier used by the write path.
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, control);
                CollectionAssert.AreEqual(new[] { "8", "10", "12" }, control.Choices.ToArray());
                Assert.AreEqual("10", control.Value);
                InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Size, "12");
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, control);
                Assert.AreEqual("12", control.Value);
                Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Size, "14"));
            }
        }

        [TestMethod]
        public void BrokenComboTraceCannotChangeAnEmptyReadOrSuppressNativeOwnershipRefusal()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            using (new VbeInspectionTrace(_ => { throw new System.IO.IOException("PRIVATE_TRACE_DESTINATION"); }).Enter())
            {
                var size = new VbeDebugWindows.OptionsControl { Name = "Size" };
                InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Size, size);
                Assert.AreEqual("10", size.Value); Assert.AreEqual(0, size.Choices.Count);
                Assert.AreEqual(IntPtr.Zero, OptionsFixtureInteger(fixture.Size, 0x157, IntPtr.Zero, IntPtr.Zero));
                var failure = Assert.ThrowsException<InvalidOperationException>(() =>
                    InvokeOptionsMethod(null, "ReadOptionsCombo", IntPtr.Zero, size));
                Assert.AreEqual("The native options ComboBox does not belong to this process.", failure.Message);
            }
        }

        [TestMethod]
        public void OptionsTraceDistinguishesUiaSizeFallbackWithoutLoggingItsChoice()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var size = fixture.Root.Add(new AutomationNode { Name = "Size", Kind = ControlType.ComboBox }.With(SelectionPattern.Pattern));
                size.Add(new AutomationNode { Name = "PRIVATE_SIZE_CHOICE", Kind = ControlType.ListItem, Selected = true }.With(SelectionItemPattern.Pattern));
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe");
                probe.Tabs(fixture.Host.Handle);
                var rows = new List<string>();
                using (new VbeInspectionTrace(rows.Add).Enter())
                    CollectionAssert.AreEqual(new[] { "PRIVATE_SIZE_CHOICE" }, probe.Controls(fixture.Host.Handle, 0).Single(x => x.Name == "Size").Choices.ToArray());
                var row = rows.Select(text => new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(text))
                    .Single(x => Equals(x["Role"], "Size"));
                Assert.AreEqual("UiAutomationCombo", row["Reader"]);
                var native = (Dictionary<string, object>)row["Native"];
                Assert.IsNull(native["CountBefore"]);
                Assert.AreEqual(false, native["ExpansionAttempted"]);
                Assert.IsFalse(string.Join("", rows).Contains("PRIVATE_SIZE_CHOICE"));
            }
        }

        /// <summary>Les fautes injectées sur les réponses entières conservent de vrais HWND détenus et ne provoquent aucune lecture de buffer trop petit.</summary>
        [TestMethod]
        public void OwnedNativeOptionsComboGuardsRejectMalformedNativeReadResultsAndSelectionFailure()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var saved = VbeDebugWindows.SendMessageInt;
                try
                {
                    foreach (int scenario in Enumerable.Range(0, 9))
                    {
                        VbeDebugWindows.SendMessageInt = (window, message, argument, value) =>
                        {
                            if (window != fixture.Font) return saved(window, message, argument, value);
                            if (message == 0x146 && scenario < 2) return new IntPtr(scenario == 0 ? -1 : 2001);
                            if (message == 0x149 && argument == IntPtr.Zero && scenario >= 2 && scenario <= 4)
                                return new IntPtr(scenario == 2 ? -1 : scenario == 3 ? 4097 : saved(window, message, argument, value).ToInt32() + 1);
                            if (message == 0x147 && scenario == 5) return new IntPtr(99);
                            if (message == 0x000E && scenario >= 6) return new IntPtr(scenario == 6 ? -1 : scenario == 7 ? 4097 : saved(window, message, argument, value).ToInt32() + 1);
                            return saved(window, message, argument, value);
                        };
                        Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Font, new VbeDebugWindows.OptionsControl()), "native read scenario " + scenario);
                    }
                    VbeDebugWindows.SendMessageInt = (window, message, argument, value) => window == fixture.Font && message == 0x14e ? new IntPtr(-1) : saved(window, message, argument, value);
                    Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Font, "Courier New"));
                    Assert.IsFalse(fixture.Notifications.Any(x => x.Item1 == 510 && (x.Item2 == 1 || x.Item2 == 9)));
                }
                finally { VbeDebugWindows.SendMessageInt = saved; }
            }
        }

        /// <summary>Les gardes refusent handles nuls, mauvaise classe, PID étranger et style owner-data non lisible.</summary>
        [TestMethod]
        public void OwnedNativeOptionsComboIdentityAndStringsGuardsStayFailClosed()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "ReadOptionsCombo", IntPtr.Zero, new VbeDebugWindows.OptionsControl()));
                Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.List, new VbeDebugWindows.OptionsControl()));
                var savedPid = VbeDebugWindows.GetWindowThreadProcessId;
                try
                {
                    VbeDebugWindows.GetWindowThreadProcessId = (IntPtr window, out uint processId) =>
                    { var result = savedPid(window, out processId); if (window == fixture.Font) processId++; return result; };
                    Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Font, new VbeDebugWindows.OptionsControl()));
                    VbeDebugWindows.GetWindowThreadProcessId = (IntPtr window, out uint processId) =>
                    { var result = OwnedOptionsFixtureProcess(window, out processId); if (window == fixture.Host.Handle) processId++; return result; };
                    var parentFailure = Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "WriteOptionsCombo", fixture.Font, "Courier New"));
                    StringAssert.Contains(parentFailure.Message, "parent does not belong to this process");
                }
                finally { VbeDebugWindows.GetWindowThreadProcessId = savedPid; }
                int original = OptionsFixtureGetStyle(fixture.Font, -16);
                try
                {
                    fixture.Host.Invoke(form => OptionsFixtureSetStyle(fixture.Font, -16, original & ~0x200));
                    Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "ReadOptionsCombo", fixture.Font, new VbeDebugWindows.OptionsControl()));
                }
                finally { fixture.Host.Invoke(form => OptionsFixtureSetStyle(fixture.Font, -16, original)); }
                Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "GuardOptionsOwnedWindow", fixture.Host.Handle, fixture.Font, "ListBox"));
                Assert.ThrowsException<InvalidOperationException>(() => InvokeOptionsMethod(null, "GuardOptionsOwnedWindow", fixture.Host.Handle, fixture.Host.Handle, "ComboBox"));
            }
        }

        /// <summary>Le propriétaire simulé ne met à jour ses palettes qu'après le vrai WM_COMMAND, puis chaque catégorie est lue et restaurée.</summary>
        [TestMethod]
        public void OwnedNativeOptionsCategoriesNotifyParentCaptureAllPalettesAndRestoreSelection()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(fixture.Host.Handle);
                var categories = (VbeDebugWindows.IFormatCategoriesOptionsProbe)probe;
                categories.SelectFormatCategory(fixture.Host.Handle, 0, "Comment");
                Assert.AreEqual("Comment", fixture.CurrentCategory);
                Assert.IsTrue(fixture.Notifications.Any(x => x.Item1 == 4905 && x.Item2 == 1 && x.Item3 == fixture.List));
                var snapshot = categories.FormatCategories(fixture.Host.Handle, 0);
                Assert.AreEqual(3, snapshot.Count); Assert.AreEqual("Comment", fixture.CurrentCategory);
                Assert.AreEqual("Automatic", snapshot[0].Palettes[0].Value);
                Assert.AreEqual("NativeIndex:1", snapshot[1].Palettes[0].Value);
                Assert.AreEqual("NativeIndex:2", snapshot[2].Palettes[0].Value);
                Assert.IsFalse(snapshot[2].Palettes[2].Enabled, "Disabled native indicator remains readable and versioned.");
                Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Missing"));
                fixture.OnCategoryNotification = name => fixture.Palettes[0].Password = name == "Keyword";
                Assert.ThrowsException<InvalidOperationException>(() => categories.FormatCategories(fixture.Host.Handle, 0));
                Assert.AreEqual("Comment", fixture.CurrentCategory, "Native palette read failure restores the initial category.");
                Assert.IsFalse(fixture.Palettes[0].Password);
            }
        }

        [TestMethod]
        public void FailedCategorySelectionRetainsRequestedObservedAndNativeIndexWithoutReplay()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(fixture.Host.Handle);
                var categories = (VbeDebugWindows.IFormatCategoriesOptionsProbe)probe;
                var requested = fixture.CategoryItems[1];
                requested.SelectedAction = () => {
                    requested.Selected = false; fixture.CategoryItems[0].Selected = true;
                };
                var failure = Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Comment"));
                StringAssert.Contains(failure.Message, "Requested=Comment; Observed=Normal; NativeIndex=0; RequestedIndex=1.");
                Assert.AreEqual(1, requested.SelectionCount, "An uncertain selection must not be replayed.");
                Assert.AreEqual(1, fixture.Notifications.Count(x => x.Item1 == 4905 && x.Item2 == 1));
                Assert.AreEqual("Normal", fixture.CurrentCategory);
            }
        }

        /// <summary>Les identités/patterns/sélections invalides des catégories et onglets échouent avant mutation des palettes.</summary>
        [TestMethod]
        public void OwnedNativeOptionsCategoryAndTabMatricesRefuseUnreadableOrAmbiguousProviders()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(fixture.Host.Handle);
                var categories = (VbeDebugWindows.IFormatCategoriesOptionsProbe)probe;
                foreach (int index in new[] { -1, 1 }) Assert.ThrowsException<InvalidOperationException>(() => probe.Controls(fixture.Host.Handle, index));
                fixture.Tab.ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id + 1;
                Assert.ThrowsException<InvalidOperationException>(() => probe.Controls(fixture.Host.Handle, 0)); fixture.Tab.ProcessId = null;
                fixture.Tab.Patterns.Clear(); Assert.ThrowsException<InvalidOperationException>(() => probe.Controls(fixture.Host.Handle, 0)); fixture.Tab.With(SelectionItemPattern.Pattern);
                fixture.Tab.Selected = false; fixture.Tab.SelectedAction = () => fixture.Tab.Selected = false;
                Assert.ThrowsException<InvalidOperationException>(() => probe.Controls(fixture.Host.Handle, 0)); fixture.Tab.SelectedAction = null;
                var original = fixture.CategoryItems[0]; var other = fixture.CategoryItems[1];
                foreach (int scenario in Enumerable.Range(0, 7))
                {
                    if (scenario == 0) fixture.Categories.Patterns.Clear();
                    if (scenario == 1) fixture.Categories.Enabled = false;
                    if (scenario == 2) original.Selected = false;
                    if (scenario == 3) other.Selected = true;
                    if (scenario == 4) other.Name = original.Name;
                    if (scenario == 5) other.Patterns.Clear();
                    if (scenario == 6) other.ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id + 1;
                    Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Comment"), "category scenario " + scenario);
                    fixture.Categories.With(SelectionPattern.Pattern); fixture.Categories.Enabled = true;
                    original.Selected = true; other.Selected = false; other.Name = "Comment"; other.With(SelectionItemPattern.Pattern); other.ProcessId = null;
                }
                fixture.Root.Children.Remove(fixture.Categories);
                Assert.AreEqual(0, categories.FormatCategories(fixture.Host.Handle, 0).Count);
                Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Normal"));
                fixture.Root.Add(fixture.Categories);
                var duplicate = fixture.Root.Add(new AutomationNode { Name = "Code Colors", Kind = ControlType.List, NativeHandle = fixture.List.ToInt32() });
                Assert.ThrowsException<InvalidOperationException>(() => categories.FormatCategories(fixture.Host.Handle, 0)); fixture.Root.Children.Remove(duplicate);
                var hidden = fixture.Palettes[0]; hidden.Name = "Font";
                Assert.ThrowsException<InvalidOperationException>(() => categories.FormatCategories(fixture.Host.Handle, 0)); hidden.Name = "Foreground";
                var extras = Enumerable.Range(3, 30).Select(i => fixture.Categories.Add(new AutomationNode { Name = "Extra " + i }.With(SelectionItemPattern.Pattern))).ToArray();
                Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Normal"));
                foreach (var extra in extras) fixture.Categories.Children.Remove(extra);
                other.Name = " "; Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Normal")); other.Name = "Comment";
                var items = fixture.Categories.Children.ToArray(); fixture.Categories.Children.Clear();
                Assert.ThrowsException<InvalidOperationException>(() => categories.SelectFormatCategory(fixture.Host.Handle, 0, "Normal"));
                foreach (var item in items) fixture.Categories.Add(item);
            }
        }

        /// <summary>Les palettes sont refusées si leur identité UIA diverge du handle réel, puis la catégorie initiale est restaurée.</summary>
        [TestMethod]
        public void OwnedNativeOptionsPaletteIdentityMatrixRestoresOnEveryFailure()
        {
            using (var fixture = new OwnedNativeOptionsControls())
            {
                var probe = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); probe.Tabs(fixture.Host.Handle);
                var categories = (VbeDebugWindows.IFormatCategoriesOptionsProbe)probe;
                var palette = fixture.Palettes[0]; int handle = palette.NativeHandle.Value;
                foreach (int scenario in Enumerable.Range(0, 4))
                {
                    fixture.OnCategoryNotification = name =>
                    {
                        bool fail = name == "Keyword";
                        palette.Offscreen = scenario == 0 && fail;
                        palette.Password = scenario == 1 && fail;
                        palette.Kind = scenario == 2 && fail ? ControlType.Edit : ControlType.ComboBox;
                        palette.NativeHandle = scenario == 3 && fail ? fixture.Palettes[1].NativeHandle : handle;
                    };
                    Assert.ThrowsException<InvalidOperationException>(() => categories.FormatCategories(fixture.Host.Handle, 0), "palette identity " + scenario);
                    Assert.AreEqual("Normal", fixture.CurrentCategory);
                    Assert.AreEqual(handle, palette.NativeHandle.Value); Assert.IsFalse(palette.Offscreen); Assert.IsFalse(palette.Password);
                    Assert.AreEqual(ControlType.ComboBox, palette.Kind);
                }
            }
        }
    }
}
