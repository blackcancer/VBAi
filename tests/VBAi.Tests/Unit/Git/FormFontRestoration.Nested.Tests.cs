using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class FormFontRestorationNestedTests
    {
        [TestMethod]
        public void ActualAttachedStateIsClonedThenOnlyDetachedDescriptorIsLoadedAndAssignedAfterFreshGuard()
        {
            var native = new Transfer(); object owner = new object(); byte[] target = Descriptor(82700);
            byte[] original = (byte[])native.Original.Data.Clone();
            int guards = 0;
            FormFontRestoration.AssignNestedCloneCore(owner, target,
                () => { guards++; native.Events.Add("guard-" + guards); }, native);
            CollectionAssert.AreEqual(new[]
            {
                "guard-1", "read", "supports-original", "class-original", "guard-2", "clone",
                "identity", "supports-clone", "class-clone", "guard-3", "load", "guard-4", "assign",
                "release-clone", "release-original"
            }, native.Events.ToArray());
            Assert.AreEqual(4, guards);
            Assert.AreSame(owner, native.ReadOwner); Assert.AreSame(owner, native.AssignedOwner);
            Assert.AreSame(native.Copy, native.AssignedFont);
            CollectionAssert.AreEqual(target, native.Copy.Data);
            CollectionAssert.AreEqual(original, native.Original.Data, "Original attached font must never be loaded.");
            Assert.AreEqual(1, native.LoadCalls); Assert.AreEqual(1, native.CloneCalls); Assert.AreEqual(1, native.AssignCalls);
            Assert.AreEqual(1, native.Original.Releases); Assert.AreEqual(1, native.Copy.Releases);
        }

        [TestMethod]
        public void InvalidDescriptorAndMissingRequiredGuardRefuseBeforeAnyNativeAcquisition()
        {
            var native = new Transfer(); int guards = 0;
            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), new byte[] { 1 }, () => guards++, native));
            Assert.ThrowsException<ArgumentNullException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), null, native));
            Assert.ThrowsException<ArgumentNullException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => guards++, null));
            Assert.AreEqual(0, guards); Assert.AreEqual(0, native.Events.Count);
        }

        [DataTestMethod]
        [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)]
        public void ChangedIdentityAtEachStageStopsWithoutExtraCloneLoadOrSetter(int rejectedGuard)
        {
            var native = new Transfer(); var failure = new InvalidOperationException("Identity changed."); int guards = 0;
            var caught = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { if (++guards == rejectedGuard) throw failure; }, native));
            Assert.AreSame(failure, caught);
            Assert.AreEqual(rejectedGuard, guards);
            Assert.AreEqual(rejectedGuard > 2 ? 1 : 0, native.CloneCalls);
            Assert.AreEqual(rejectedGuard > 3 ? 1 : 0, native.LoadCalls);
            Assert.AreEqual(0, native.AssignCalls);
            Assert.AreEqual(rejectedGuard > 1 ? 1 : 0, native.Original.Releases);
            Assert.AreEqual(rejectedGuard > 2 ? 1 : 0, native.Copy.Releases);
            CollectionAssert.AreEqual(Descriptor(82500), native.Original.Data);
        }

        [DataTestMethod]
        [DataRow("null-original")] [DataRow("original-interface")] [DataRow("original-class")]
        [DataRow("null-clone")] [DataRow("alias-clone")] [DataRow("same-wrapper")]
        [DataRow("clone-interface")] [DataRow("clone-class")]
        public void UnsupportedMissingOrAliasedFontsNeverLoadOrAssign(string condition)
        {
            var native = new Transfer();
            if (condition == "null-original") native.NoOriginal = true;
            else if (condition == "original-interface") native.Original.Supported = false;
            else if (condition == "original-class") native.Original.Class = Guid.NewGuid();
            else if (condition == "null-clone") native.NoClone = true;
            else if (condition == "alias-clone") native.Copy.Identity = native.Original.Identity;
            else if (condition == "same-wrapper") native.SameWrapper = true;
            else if (condition == "clone-interface") native.Copy.Supported = false;
            else if (condition == "clone-class") native.Copy.Class = Guid.NewGuid();
            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { }, native));
            Assert.AreEqual(0, native.LoadCalls); Assert.AreEqual(0, native.AssignCalls);
            Assert.AreEqual(condition.StartsWith("original-") || condition == "null-original" ? 0 : 1, native.CloneCalls);
            Assert.AreEqual(condition == "null-original" ? 0 : condition == "same-wrapper" ? 2 : 1, native.Original.Releases);
            Assert.AreEqual(native.CloneCalls == 1 && !native.NoClone && !native.SameWrapper ? 1 : 0, native.Copy.Releases);
            CollectionAssert.AreEqual(Descriptor(82500), native.Original.Data);
        }

        [DataTestMethod]
        [DataRow("read", "MSForms.Font.get")]
        [DataRow("supports-original", "AttachedFont.interfaces")]
        [DataRow("class-original", "AttachedFont.IPersistStream.GetClassID")]
        [DataRow("clone", "IFont.Clone")]
        [DataRow("identity", "ClonedFont.IUnknown.identity")]
        [DataRow("supports-clone", "ClonedFont.interfaces")]
        [DataRow("class-clone", "ClonedFont.IPersistStream.GetClassID")]
        [DataRow("load", "IPersistStream.Load")]
        [DataRow("assign", "MSForms.Font.set")]
        public void NativeStageFailureKeepsOperationHResultAndAcquisitionProvenanceWithoutRetry(string stage, string operation)
        {
            var failure = new COMException("Synthetic native failure.", unchecked((int)0x80020003));
            var native = new Transfer { FailedAt = stage, Failure = failure };
            var caught = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { }, native));
            Assert.AreSame(failure, caught.InnerException); StringAssert.Contains(caught.Message, operation);
            StringAssert.Contains(caught.Message, "80020003");
            Assert.AreEqual(1, native.Events.Count(value => value == stage));
            Assert.AreEqual(stage == "read" ? 0 : 1, native.Original.Releases);
            bool acquiredClone = new[] { "identity", "supports-clone", "class-clone", "load", "assign" }.Contains(stage);
            Assert.AreEqual(acquiredClone ? 1 : 0, native.Copy.Releases);
            Assert.IsTrue(native.CloneCalls <= 1); Assert.IsTrue(native.LoadCalls <= 1); Assert.IsTrue(native.AssignCalls <= 1);
            Assert.AreEqual(stage == "assign" ? 1 : 0, native.AssignCalls);
            CollectionAssert.AreEqual(Descriptor(82500), native.Original.Data);
        }

        [TestMethod]
        public void UnsupportedCloneOperationNeverUsesFactoryFallbackAndStillReleasesOriginal()
        {
            var failure = new NotSupportedException("Clone is not implemented.");
            var native = new Transfer { FailedAt = "clone", Failure = failure };
            var caught = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { }, native));
            Assert.AreSame(failure, caught.InnerException); StringAssert.Contains(caught.Message, "IFont.Clone");
            Assert.AreEqual(1, native.CloneCalls); Assert.AreEqual(0, native.LoadCalls); Assert.AreEqual(0, native.AssignCalls);
            Assert.AreEqual(1, native.Original.Releases); Assert.AreEqual(0, native.Copy.Releases);
        }

        [TestMethod]
        public void UncertainSetterFailureRemainsPrimaryBeforeEveryReleaseFailure()
        {
            var failure = new COMException("Uncertain setter outcome.", unchecked((int)0x80004005));
            var cloneCleanup = new InvalidOperationException("Clone release failed.");
            var originalCleanup = new InvalidOperationException("Original release failed.");
            var native = new Transfer { FailedAt = "assign", Failure = failure };
            native.Copy.ReleaseFailure = cloneCleanup; native.Original.ReleaseFailure = originalCleanup;
            var caught = Assert.ThrowsException<AggregateException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { }, native));
            Assert.AreEqual(3, caught.InnerExceptions.Count);
            Assert.AreSame(failure, caught.InnerExceptions[0].InnerException);
            Assert.AreSame(cloneCleanup, caught.InnerExceptions[1]); Assert.AreSame(originalCleanup, caught.InnerExceptions[2]);
            Assert.AreEqual(1, native.AssignCalls); Assert.AreEqual(1, native.LoadCalls); Assert.AreEqual(1, native.CloneCalls);
            Assert.AreEqual(1, native.Copy.Releases); Assert.AreEqual(1, native.Original.Releases);
        }

        [TestMethod]
        public void SuccessfulTransferWithReleaseFailureCannotBeReportedAsCleanSuccess()
        {
            var native = new Transfer(); var cleanup = new InvalidOperationException("Clone release failed.");
            native.Copy.ReleaseFailure = cleanup;
            var caught = Assert.ThrowsException<AggregateException>(() => FormFontRestoration.AssignNestedCloneCore(
                new object(), Descriptor(82700), () => { }, native));
            Assert.AreSame(cleanup, caught.InnerExceptions.Single()); Assert.AreEqual(1, native.AssignCalls);
            Assert.AreEqual(1, native.Copy.Releases); Assert.AreEqual(1, native.Original.Releases);
        }

        [TestMethod]
        public void NativeCloneContractMatchesSdkVtableCurrencyBoolAndInterfaceOwnership()
        {
            Type type = typeof(FormFontRestoration).GetNestedType("NestedNativeFont", BindingFlags.NonPublic);
            Assert.AreEqual(new Guid("BEF6E002-A874-101A-8BBA-00AA00300CAB"), type.GUID);
            Assert.IsTrue(type.IsImport);
            CollectionAssert.AreEqual(new[] { "GetName", "SetName", "GetSize", "SetSize", "GetBold", "SetBold",
                "GetItalic", "SetItalic", "GetUnderline", "SetUnderline", "GetStrikethrough", "SetStrikethrough",
                "GetWeight", "SetWeight", "GetCharset", "SetCharset", "GetHFont", "Clone" },
                type.GetMethods().OrderBy(method => method.MetadataToken).Select(method => method.Name).ToArray());
            Assert.AreEqual(UnmanagedType.Currency,
                type.GetMethod("GetSize").GetParameters()[0].GetCustomAttribute<MarshalAsAttribute>().Value);
            Assert.AreEqual(UnmanagedType.Bool,
                type.GetMethod("GetBold").GetParameters()[0].GetCustomAttribute<MarshalAsAttribute>().Value);
            ParameterInfo clone = type.GetMethod("Clone").GetParameters()[0];
            Assert.IsTrue(clone.IsOut); Assert.AreEqual(type.MakeByRefType(), clone.ParameterType);
            Assert.AreEqual(UnmanagedType.Interface, clone.GetCustomAttribute<MarshalAsAttribute>().Value);
        }

        [TestMethod]
        public void ProductionEntryRefusesMtaBeforeOwnerAcquisitionOrIdentityCallbacks()
        {
            int guards = 0; var owner = new UnreadableOwner(); Exception observed = null;
            var thread = new Thread(() =>
            {
                try
                {
                    typeof(FormFontRestoration).GetMethod("AssignNestedClone", BindingFlags.NonPublic | BindingFlags.Static)
                        .Invoke(null, new object[] { owner, Descriptor(82700), new Action(() => guards++) });
                }
                catch (TargetInvocationException error) { observed = error.InnerException; }
            });
            thread.SetApartmentState(ApartmentState.MTA); thread.Start();
            Assert.IsTrue(thread.Join(5000), "The entry refusal must not start a native operation.");
            Assert.IsInstanceOfType(observed, typeof(InvalidOperationException));
            StringAssert.Contains(observed.Message, "owning STA");
            Assert.AreEqual(0, guards); Assert.AreEqual(0, owner.Reads);
        }

        public sealed class UnreadableOwner
        {
            public int Reads { get; private set; }
            public object Font { get { Reads++; throw new InvalidOperationException("Owner getter must not run."); } }
        }

        private static byte[] Descriptor(uint size)
        {
            byte[] data = { 1, 0, 0, 0, 144, 1, 0, 0, 0, 0, 6, 84, 97, 104, 111, 109, 97 };
            Array.Copy(BitConverter.GetBytes(size), 0, data, 6, 4); return data;
        }

        private sealed class Font
        {
            internal object Identity = new object();
            internal Guid Class = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851");
            internal bool Supported = true;
            internal byte[] Data = Descriptor(82500);
            internal int Releases;
            internal Exception ReleaseFailure;
        }

        private sealed class Transfer : FormFontRestoration.INestedFontTransfer
        {
            internal readonly Font Original = new Font(), Copy = new Font();
            internal readonly List<string> Events = new List<string>();
            internal bool NoOriginal, NoClone, SameWrapper;
            internal string FailedAt;
            internal Exception Failure;
            internal int CloneCalls, LoadCalls, AssignCalls;
            internal object ReadOwner, AssignedOwner, AssignedFont;
            private void Visit(string stage) { Events.Add(stage); if (stage == FailedAt) throw Failure; }
            private string Label(object font) { return ReferenceEquals(font, Original) ? "original" : "clone"; }
            public object ReadFont(object owner) { ReadOwner = owner; Visit("read"); return NoOriginal ? null : Original; }
            public bool SupportsTransfer(object font) { Visit("supports-" + Label(font)); return ((Font)font).Supported; }
            public Guid ClassId(object font) { Visit("class-" + Label(font)); return ((Font)font).Class; }
            public object Clone(object font)
            {
                CloneCalls++; Visit("clone"); Assert.AreSame(Original, font);
                Copy.Data = (byte[])Original.Data.Clone(); return NoClone ? null : SameWrapper ? Original : Copy;
            }
            public bool SameIdentity(object first, object second)
            { Visit("identity"); return ReferenceEquals(((Font)first).Identity, ((Font)second).Identity); }
            public void Load(object font, byte[] data)
            { LoadCalls++; Visit("load"); Assert.AreSame(Copy, font); Copy.Data = (byte[])data.Clone(); }
            public void Assign(object owner, object font)
            { AssignCalls++; Visit("assign"); AssignedOwner = owner; AssignedFont = font; }
            public void Release(object font)
            {
                var value = (Font)font; value.Releases++; Visit("release-" + Label(font));
                if (value.ReleaseFailure != null) throw value.ReleaseFailure;
            }
        }
    }
}
