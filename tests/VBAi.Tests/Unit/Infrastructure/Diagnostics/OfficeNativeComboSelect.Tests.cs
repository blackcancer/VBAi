using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeNativeComboSelectTests
    {
        private sealed class PassiveForm : Form
        {
            protected override bool ShowWithoutActivation => true;
        }
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

        [TestMethod]
        public void ExactSelectionNotifiesOnceOnlyAfterReadbackAndDoesNotReplay()
        {
            var actions = new List<string>();
            OfficeNativeComboSelect.SelectOnce(() => actions.Add("guard"), 1,
                () => { actions.Add("set"); return 1; }, () => { actions.Add("read"); return "After"; },
                "After", () => actions.Add("notify"), actions.Add);
            CollectionAssert.AreEqual(new[] { "guard", "NativeComboSetIntentOnce", "guard", "set", "guard", "read",
                "guard", "NativeComboNotifyIntentOnce", "guard", "notify", "guard", "read", "NativeComboSelectionReadback" }, actions);
        }

        [DataTestMethod]
        [DataRow("setThrow", 1, 0, 0)][DataRow("setRejected", 1, 0, 0)]
        [DataRow("firstRead", 1, 1, 0)][DataRow("notifyThrow", 1, 1, 1)]
        [DataRow("secondRead", 1, 2, 1)][DataRow("setReceipt", 0, 0, 0)]
        [DataRow("notifyReceipt", 1, 1, 0)]
        public void FailureStopsTheSelectionWithoutFallbackOrRepeatedMutation(string stage, int sets, int reads, int notifications)
        {
            int set = 0, read = 0, notify = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboSelect.SelectOnce(() => { }, 1,
                () => { set++; if (stage == "setThrow") throw new InvalidOperationException(); return stage == "setRejected" ? 0 : 1; },
                () => { read++; return stage == "firstRead" || stage == "secondRead" && read == 2 ? "Changed" : "After"; },
                "After", () => { notify++; if (stage == "notifyThrow") throw new InvalidOperationException(); },
                phase => { if (stage == "setReceipt" && phase == "NativeComboSetIntentOnce" || stage == "notifyReceipt" && phase == "NativeComboNotifyIntentOnce") throw new InvalidOperationException(); }));
            Assert.AreEqual(sets, set); Assert.AreEqual(reads, read); Assert.AreEqual(notifications, notify);
        }

        [DataTestMethod]
        [DataRow(1, 0, 0)][DataRow(2, 0, 0)][DataRow(3, 1, 0)]
        [DataRow(4, 1, 0)][DataRow(5, 1, 0)][DataRow(6, 1, 1)]
        public void OwnershipLossAtEveryBoundaryPreventsRemainingActions(int boundary, int expectedSets, int expectedNotifications)
        {
            int guards = 0, sets = 0, notifications = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboSelect.SelectOnce(
                () => { if (++guards == boundary) throw new InvalidOperationException("Original owner ended."); }, 1,
                () => { sets++; return 1; }, () => "After", "After", () => notifications++, unused => { }));
            Assert.AreEqual(expectedSets, sets); Assert.AreEqual(expectedNotifications, notifications);
        }

        [DataTestMethod][DataRow(-1L)][DataRow(4096L)]
        public void InvalidChoiceCannotDispatchSelection(long index)
        {
            Assert.ThrowsException<ArgumentException>(() => OfficeNativeComboSelect.SelectOnce(
                () => Assert.Fail(), index, () => { Assert.Fail(); return 0; }, () => "After", "After", () => Assert.Fail(), unused => Assert.Fail()));
        }

        [DataTestMethod][DataRow(DrawMode.Normal)][DataRow(DrawMode.OwnerDrawFixed)]
        public void ActualWinFormsSelectionRaisesOneEventWithoutOpeningPopupOrChangingFocus(DrawMode drawMode)
        {
            OnSta(() =>
            {
                using (var form = new PassiveForm { ShowInTaskbar = false })
                using (var input = new TextBox { Top = 40 })
                using (var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, DrawMode = drawMode })
                {
                    combo.Items.AddRange(new object[] { "Before", "After" }); combo.SelectedIndex = 0;
                    form.Controls.Add(combo); form.Controls.Add(input); form.ActiveControl = input; form.Show(); Application.DoEvents();
                    int changes = 0; combo.SelectedIndexChanged += (sender, args) => { changes++; Assert.AreEqual("After", combo.SelectedItem); };
                    uint pid; uint thread = GetWindowThreadProcessId(combo.Handle, out pid);
                    OfficeNativeComboSelect.Select(combo.Handle, form.Handle, (int)pid, thread, "After", () => { }, unused => { });
                    Assert.AreEqual(1, combo.SelectedIndex); Assert.AreEqual("After", combo.Text);
                    Assert.AreEqual(1, changes, "The native parent notification must reach the real WinForms handler exactly once.");
                    Assert.AreSame(input, form.ActiveControl); Assert.IsFalse(combo.DroppedDown);
                    form.Close();
                }
            });
        }

        [DataTestMethod][DataRow("duplicate")][DataRow("missing")][DataRow("disabled")]
        [DataRow("process")][DataRow("thread")][DataRow("owner")]
        public void ActualSelectorRejectsAmbiguityAndChangedOwnershipBeforeMutation(string fault)
        {
            OnSta(() =>
            {
                using (var form = new PassiveForm { ShowInTaskbar = false })
                using (var foreign = new PassiveForm { ShowInTaskbar = false })
                using (var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList })
                {
                    combo.Items.AddRange(new object[] { "Before", "After" });
                    if (fault == "duplicate") combo.Items.Add("After"); combo.SelectedIndex = 0;
                    form.Controls.Add(combo); form.Show(); Application.DoEvents();
                    int changes = 0; combo.SelectedIndexChanged += (sender, args) => changes++;
                    uint pid; uint thread = GetWindowThreadProcessId(combo.Handle, out pid);
                    if (fault == "disabled") combo.Enabled = false;
                    Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboSelect.Select(combo.Handle,
                        fault == "owner" ? foreign.Handle : form.Handle, fault == "process" ? (int)pid + 1 : (int)pid,
                        fault == "thread" ? thread + 1 : thread, fault == "missing" ? "Absent" : "After", () => { }, unused => { }));
                    Assert.AreEqual(0, combo.SelectedIndex); Assert.AreEqual(0, changes); Assert.IsFalse(combo.DroppedDown); form.Close();
                }
            });
        }

        [DataTestMethod][DataRow(false)][DataRow(true)]
        public void ActualSelectorReadbackSurvivesTemporaryDisableCausedByTheSingleNotification(bool disableParent)
        {
            OnSta(() =>
            {
                using (var form = new PassiveForm { ShowInTaskbar = false })
                using (var parent = new Panel { Dock = DockStyle.Fill })
                using (var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList })
                {
                    combo.Items.AddRange(new object[] { "Before", "After" }); combo.SelectedIndex = 0;
                    parent.Controls.Add(combo); form.Controls.Add(parent); form.Show(); Application.DoEvents();
                    int changes = 0;
                    combo.SelectedIndexChanged += (sender, args) =>
                    {
                        changes++;
                        if (disableParent) parent.Enabled = false; else combo.Enabled = false;
                    };
                    uint pid; uint thread = GetWindowThreadProcessId(combo.Handle, out pid);
                    var receipts = new List<string>();
                    OfficeNativeComboSelect.Select(combo.Handle, form.Handle, (int)pid, thread, "After", () => { }, receipts.Add);
                    Assert.AreEqual(1, combo.SelectedIndex); Assert.AreEqual("After", combo.Text);
                    Assert.AreEqual(1, changes); Assert.IsFalse(combo.Enabled); Assert.IsFalse(combo.DroppedDown);
                    CollectionAssert.AreEqual(new[] { "NativeComboSetIntentOnce", "NativeComboNotifyIntentOnce", "NativeComboSelectionReadback" }, receipts);
                    form.Close();
                }
            });
        }

        [DataTestMethod][DataRow("NativeComboSetIntentOnce", 0)][DataRow("NativeComboNotifyIntentOnce", 1)]
        public void ActualSelectorRejectsDisabledControlAtEachMutationBoundary(string phase, int expectedIndex)
        {
            OnSta(() =>
            {
                using (var form = new PassiveForm { ShowInTaskbar = false })
                using (var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList })
                {
                    combo.Items.AddRange(new object[] { "Before", "After" }); combo.SelectedIndex = 0;
                    form.Controls.Add(combo); form.Show(); Application.DoEvents();
                    int changes = 0; combo.SelectedIndexChanged += (sender, args) => changes++;
                    uint pid; uint thread = GetWindowThreadProcessId(combo.Handle, out pid);
                    Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboSelect.Select(combo.Handle, form.Handle,
                        (int)pid, thread, "After", () => { }, observed => { if (observed == phase) combo.Enabled = false; }));
                    Assert.AreEqual(expectedIndex, combo.SelectedIndex); Assert.AreEqual(0, changes);
                    Assert.IsFalse(combo.DroppedDown); form.Close();
                }
            });
        }

        private static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(15)), "The owned selector test must finish within its bound.");
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
