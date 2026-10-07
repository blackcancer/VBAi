using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using Candidate = VBAi.VbeProjectComponents.AccessSaveConfirmationCandidate;
using Component = VBAi.VbeProjectComponents.AccessSaveApprovedComponent;
using Control = VBAi.VbeProjectComponents.AccessSaveDialogControl;
using Dialog = VBAi.VbeProjectComponents.AccessSaveDialogSnapshot;
using Gate = VBAi.VbeProjectComponents.AccessSaveConfirmation;
using Inventory = VBAi.VbeProjectComponents.AccessSaveDialogInventory;
using Item = VBAi.VbeProjectComponents.AccessSaveDialogItem;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the exact approved prompt and single delivery with pure native snapshots.</summary>
    [TestClass]
    public sealed class AccessSaveConfirmationTests
    {
        private sealed class Fixture
        {
            internal Inventory Current = Empty();
            internal int Reads, Enqueues, OwnerChecks, ContextChecks;
            internal Action AfterRead;
            internal bool OwnerValid = true, PostAccepted = true, PostThrows;
            internal Gate Gate;
            internal Fixture()
            {
                Gate = new Gate(() => { Reads++; AfterRead?.Invoke(); return Current; },
                    candidate => { Enqueues++; if (PostThrows) throw new InvalidOperationException("Uncertain post"); return PostAccepted; },
                    () => { OwnerChecks++; if (!OwnerValid) throw new InvalidOperationException("Owner changed"); },
                    10, 20, new[] { new Component("ApprovedClass", 2), new Component("ApprovedModule", 1) });
            }
            internal Candidate Ready()
            {
                Gate.Prepare(); Gate.RequireBeforeSave(); Current = Known();
                return Gate.Observe();
            }
            internal void Context() { ContextChecks++; }
        }

        private static Inventory Empty() => new Inventory { Complete = true, ProcessId = 10, OwnerThreadId = 20, Dialogs = new Dialog[0] };
        private static Control Child(int id, string type, string text, int hwnd) => new Control
        {
            Window = new IntPtr(hwnd),
            Id = id,
            Class = type,
            Text = text,
            ProcessId = 10,
            ThreadId = 20,
            Visible = true,
            Enabled = true,
            Style = 0x50010000
        };
        private static Inventory Known()
        {
            var list = Child(5142, "ListBox", null, 105);
            list.Style = 0x5001016B; list.ItemsComplete = true;
            list.Items = new[] { new Item { Text = "Module: ApprovedClass", Selected = true },
                new Item { Text = "Module: ApprovedModule", Selected = true } };
            return new Inventory
            {
                Complete = true,
                ProcessId = 10,
                OwnerThreadId = 20,
                Dialogs = new[] {
                new Dialog { Window = new IntPtr(100), ProcessId = 10, ThreadId = 20, Style = 0x94C800C4,
                    Class = "#32770", Caption = "Enregistrer", Visible = true, Enabled = true, ChildrenComplete = true,
                    Controls = new[] { Child(1, "Button", "&Oui", 101), Child(7, "Button", "&Non pour tout", 102),
                        Child(2, "Button", "Annuler", 103),
                        Child(5271, "Static", "Enregistrer les modifications apportées aux objets suivants\u00A0?", 104), list,
                        new Control { Window = new IntPtr(106), Id = -1, Class = "Static", Text = "DAL=on",
                            ProcessId = 10, ThreadId = 20, Enabled = true, Visible = false } } }
            }
            };
        }
        private static Control List(Inventory inventory) => inventory.Dialogs[0].Controls.Single(control => control.Id == 5142);

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void FinalSnapshotMustCompleteBeforeTheSaveDeliveryDeadline(bool expiresDuringFinalRead)
        {
            var f = new Fixture(); var candidate = f.Ready();
            int elapsedMilliseconds = 0, deadlineChecks = 0;
            f.AfterRead = () =>
            {
                // Prepare, pre-Save and Observe are reads 1..3. Confirm freezes the
                // candidate on read 4, checks approval, then performs final read 5.
                if (f.Reads == 5) elapsedMilliseconds = expiresDuringFinalRead ? 3000 : 2999;
            };
            Action approvedContext = () =>
            {
                f.Context();
                Assert.AreEqual(0, elapsedMilliseconds, "Approval precedes the final native snapshot.");
            };
            Action deliveryDeadline = () =>
            {
                deadlineChecks++;
                Assert.AreEqual(5, f.Reads, "The delivery deadline must follow the final native snapshot.");
                if (elapsedMilliseconds >= 3000) throw new InvalidOperationException("Delivery deadline expired");
            };
            if (expiresDuringFinalRead)
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, approvedContext, deliveryDeadline));
                Assert.AreEqual(0, f.Enqueues); Assert.AreEqual(0, f.Gate.ConfirmationAttempts);
                Assert.IsFalse(f.Gate.ConfirmationQueued); Assert.IsTrue(f.Gate.ConfirmationPending);
                Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, approvedContext, deliveryDeadline));
                Assert.AreEqual(0, f.Enqueues, "Expiry consumes the confirmation call without replay.");
            }
            else
            {
                f.Gate.Confirm(candidate, approvedContext, deliveryDeadline);
                Assert.AreEqual(1, f.Enqueues); Assert.AreEqual(1, f.Gate.ConfirmationAttempts);
                Assert.IsTrue(f.Gate.ConfirmationQueued);
            }
            Assert.AreEqual(1, deadlineChecks); Assert.AreEqual(1, f.ContextChecks);
        }

        [TestMethod]
        public void ApprovedClassAndModuleQueueOnceAndWaitForOriginalModalToDisappear()
        {
            var f = new Fixture(); var candidate = f.Ready();
            Assert.AreEqual(0, f.Enqueues); Assert.IsTrue(f.Gate.ConfirmationPending);
            f.Gate.Confirm(candidate, f.Context);
            Assert.AreEqual(1, f.ContextChecks); Assert.AreEqual(1, f.Enqueues);
            Assert.AreEqual(1, f.Gate.ConfirmationAttempts); Assert.IsTrue(f.Gate.ConfirmationQueued);
            Assert.AreSame(candidate, f.Gate.Observe(), "A queued response need not synchronously dismiss the modal.");
            Assert.IsTrue(f.Gate.ConfirmationPending);
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, f.Context));
            f.Current = Empty(); Assert.IsNull(f.Gate.Observe()); Assert.IsFalse(f.Gate.ConfirmationPending);
            Assert.IsTrue(f.Gate.ConfirmationQueued); Assert.AreEqual(1, f.Enqueues);
        }

        [TestMethod]
        public void SingleApprovedSelectedObjectAlsoMatchesTheQualifiedSaveAllShape()
        {
            var f = new Fixture(); f.Gate.Prepare(); f.Gate.RequireBeforeSave(); f.Current = Known();
            List(f.Current).Items = new[] { new Item { Text = "Module: ApprovedClass", Selected = true } };
            f.Gate.Confirm(f.Gate.Observe(), f.Context);
            Assert.AreEqual(1, f.Enqueues);
        }

        [DataTestMethod]
        [DataRow("inventoryIncomplete"), DataRow("inventoryPid"), DataRow("inventoryThread"), DataRow("multipleDialogs")]
        [DataRow("dialogThread"), DataRow("dialogPid"), DataRow("dialogClass"), DataRow("caption"), DataRow("dialogDisabled")]
        [DataRow("childrenIncomplete"), DataRow("buttonId"), DataRow("buttonCaption"), DataRow("buttonClass")]
        [DataRow("buttonDisabled"), DataRow("buttonHidden"), DataRow("childThread"), DataRow("childPid")]
        [DataRow("duplicateId"), DataRow("duplicateHandle"), DataRow("unknownVisibleChild"), DataRow("prompt")]
        [DataRow("listId"), DataRow("listClass"), DataRow("listDisabled"), DataRow("listHidden")]
        [DataRow("ownerDrawNoStrings"), DataRow("noData"), DataRow("singleSelect"), DataRow("itemsIncomplete")]
        [DataRow("emptyItems"), DataRow("tooManyItems"), DataRow("unknownName"), DataRow("wrongObjectType")]
        [DataRow("notSelected"), DataRow("duplicateItem"), DataRow("longItem"), DataRow("nullItem")]
        public void UnqualifiedPromptOrObjectsNeverEnqueue(string change)
        {
            var f = new Fixture(); f.Gate.Prepare(); f.Gate.RequireBeforeSave(); f.Current = Known();
            var dialog = f.Current.Dialogs[0]; var yes = dialog.Controls[0]; var list = List(f.Current);
            switch (change)
            {
                case "inventoryIncomplete": f.Current.Complete = false; break;
                case "inventoryPid": f.Current.ProcessId++; break;
                case "inventoryThread": f.Current.OwnerThreadId++; break;
                case "multipleDialogs": f.Current.Dialogs = new[] { dialog, dialog }; break;
                case "dialogThread": dialog.ThreadId++; break;
                case "dialogPid": dialog.ProcessId++; break;
                case "dialogClass": dialog.Class = "Other"; break;
                case "caption": dialog.Caption = "Unknown"; break;
                case "dialogDisabled": dialog.Enabled = false; break;
                case "childrenIncomplete": dialog.ChildrenComplete = false; break;
                case "buttonId": yes.Id = 6; break;
                case "buttonCaption": yes.Text = "OK"; break;
                case "buttonClass": yes.Class = "Edit"; break;
                case "buttonDisabled": yes.Enabled = false; break;
                case "buttonHidden": yes.Visible = false; break;
                case "childThread": yes.ThreadId++; break;
                case "childPid": yes.ProcessId++; break;
                case "duplicateId": dialog.Controls[1].Id = 1; break;
                case "duplicateHandle": dialog.Controls[1].Window = yes.Window; break;
                case "unknownVisibleChild": dialog.Controls[5].Visible = true; break;
                case "prompt": dialog.Controls[3].Text = "Save all documents?"; break;
                case "listId": list.Id++; break;
                case "listClass": list.Class = "Edit"; break;
                case "listDisabled": list.Enabled = false; break;
                case "listHidden": list.Visible = false; break;
                case "ownerDrawNoStrings": list.Style &= ~0x40u; break;
                case "noData": list.Style |= 0x2000; break;
                case "singleSelect": list.Style &= ~(0x8u | 0x800u); break;
                case "itemsIncomplete": list.ItemsComplete = false; break;
                case "emptyItems": list.Items = new Item[0]; break;
                case "tooManyItems": list.Items = Enumerable.Range(0, 1001).Select(i => new Item { Text = "Module: ApprovedModule", Selected = true }).ToArray(); break;
                case "unknownName": list.Items[0].Text = "Module: UserProductionModule"; break;
                case "wrongObjectType": list.Items[0].Text = "Form: ApprovedClass"; break;
                case "notSelected": list.Items[0].Selected = false; break;
                case "duplicateItem": list.Items[1].Text = list.Items[0].Text; break;
                case "longItem": list.Items[0].Text = new string('X', 513); break;
                case "nullItem": list.Items[0] = null; break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Observe(), change);
            Assert.AreEqual(0, f.Enqueues, change); Assert.AreEqual(0, f.Gate.ConfirmationAttempts, change);
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Observe(), "A failed gate cannot be reused.");
        }

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void PreexistingKnownOrUnknownModalRefusesBeforeSave(bool unknown)
        {
            var f = new Fixture(); f.Current = Known();
            if (unknown) f.Current.Dialogs[0].Caption = "Unrelated dialog";
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Prepare());
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Prepare());
            Assert.AreEqual(0, f.Enqueues);
        }

        [TestMethod]
        public void ModalAppearingAfterPreparationRefusesTheImmediatePreSaveGate()
        {
            var f = new Fixture(); f.Gate.Prepare(); f.Current = Known();
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.RequireBeforeSave());
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.RequireBeforeSave());
            Assert.AreEqual(0, f.Enqueues);
        }

        [DataTestMethod]
        [DataRow("beforeContext"), DataRow("insideContext"), DataRow("contextThrows"), DataRow("ownerChanged")]
        public void CandidateOrApprovalChangesConsumeTheConfirmationCallWithoutEnqueue(string change)
        {
            var f = new Fixture(); var candidate = f.Ready();
            if (change == "beforeContext") f.Current.Dialogs[0].Style++;
            if (change == "ownerChanged") f.OwnerValid = false;
            Action context = () =>
            {
                f.Context();
                if (change == "insideContext") List(f.Current).Items[0].Selected = false;
                if (change == "contextThrows") throw new InvalidOperationException("Approval revoked");
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, context));
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, context));
            Assert.AreEqual(0, f.Enqueues); Assert.AreEqual(0, f.Gate.ConfirmationAttempts);
        }

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void RejectedOrThrowingPostRemainsUncertainAndCannotBeRetried(bool throws)
        {
            var f = new Fixture(); var candidate = f.Ready(); f.PostAccepted = false; f.PostThrows = throws;
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, f.Context));
            Assert.AreEqual(1, f.Enqueues); Assert.AreEqual(1, f.Gate.ConfirmationAttempts);
            Assert.IsFalse(f.Gate.ConfirmationQueued); Assert.IsTrue(f.Gate.ConfirmationPending);
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Confirm(candidate, f.Context));
            Assert.AreEqual(1, f.Enqueues);
        }

        [TestMethod]
        public void DifferentModalAfterQueuedConfirmationIsRefusedWithoutAnotherPost()
        {
            var f = new Fixture(); var candidate = f.Ready(); f.Gate.Confirm(candidate, f.Context);
            f.Current = Known(); f.Current.Dialogs[0].Window = new IntPtr(200);
            Assert.ThrowsException<InvalidOperationException>(() => f.Gate.Observe());
            Assert.AreEqual(1, f.Enqueues); Assert.IsTrue(f.Gate.ConfirmationQueued);
        }

        [TestMethod]
        public void CandidateCannotBeTransferredBetweenSaveGates()
        {
            var source = new Fixture(); var foreign = source.Ready();
            var other = new Fixture(); other.Ready();
            Assert.ThrowsException<InvalidOperationException>(() => other.Gate.Confirm(foreign, other.Context));
            Assert.AreEqual(0, other.Enqueues);
        }

        [DataTestMethod]
        [DataRow(0), DataRow(3), DataRow(100)]
        public void NonModuleOrClassComponentsCannotEnterApprovalSet(int type)
        {
            Assert.ThrowsException<InvalidOperationException>(() => new Gate(Empty, candidate => true, () => { },
                10, 20, new[] { new Component("ProductionForm", type) }));
        }

        [TestMethod]
        public void EmptyApprovalAllowsNoDialogMetadataSaveButCannotConfirmAnyListedObject()
        {
            Inventory current = Empty(); int enqueues = 0;
            var gate = new Gate(() => current, candidate => { enqueues++; return true; }, () => { },
                10, 20, new Component[0]);
            gate.Prepare(); gate.RequireBeforeSave(); Assert.IsNull(gate.Observe());
            current = Known();
            Assert.ThrowsException<InvalidOperationException>(() => gate.Observe());
            Assert.AreEqual(0, enqueues);
        }

        [TestMethod]
        public void ExistingLongComponentNameIsPreservedInExactApproval()
        {
            string name = new string('N', 64);
            Inventory current = Empty(); int enqueues = 0;
            var gate = new Gate(() => current, candidate => { enqueues++; return true; }, () => { },
                10, 20, new[] { new Component(name, 1) });
            gate.Prepare(); gate.RequireBeforeSave(); current = Known();
            List(current).Items = new[] { new Item { Text = "Module: " + name, Selected = true } };
            gate.Confirm(gate.Observe(), () => { });
            Assert.AreEqual(1, enqueues);
        }

        [TestMethod]
        public void DuplicateApprovalsCannotAuthorizeNativeObjects()
        {
            Assert.ThrowsException<InvalidOperationException>(() => new Gate(Empty, candidate => true, () => { }, 10, 20,
                new[] { new Component("Same", 1), new Component("Same", 2) }));
            Assert.ThrowsException<InvalidOperationException>(() => new Gate(Empty, candidate => true, () => { }, 10, 20,
                new[] { new Component("Same", 1), new Component("same", 2) }));
        }
    }
}
