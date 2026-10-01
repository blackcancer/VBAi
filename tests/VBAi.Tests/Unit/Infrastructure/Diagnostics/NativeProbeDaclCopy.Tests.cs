using System;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests
{
    /// <summary>Pure binary ACL contracts for the disposable native export probe; no filesystem or token mutation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeProbeDaclCopyTests
    {
        private static readonly SecurityIdentifier User = new SecurityIdentifier("S-1-5-21-1-2-3-1001");
        private static readonly SecurityIdentifier Other = new SecurityIdentifier("S-1-5-21-1-2-3-1002");

        [TestMethod]
        public void RawCopyPreservesGenericMasksAceOrderTypesAndAllOtherFlags()
        {
            var acl = new RawAcl(GenericAcl.AclRevisionDS, 3);
            acl.InsertAce(0, new CommonAce(AceFlags.Inherited | AceFlags.ObjectInherit, AceQualifier.AccessDenied,
                unchecked((int)0x80000000U), Other, false, null)); // GENERIC_READ, negative int
            acl.InsertAce(1, new CommonAce(AceFlags.Inherited | AceFlags.ContainerInherit | AceFlags.InheritOnly,
                AceQualifier.AccessAllowed, 0x10000000, Other, false, null)); // GENERIC_ALL
            acl.InsertAce(2, new ObjectAce(AceFlags.Inherited | AceFlags.NoPropagateInherit, AceQualifier.AccessAllowed,
                0x10000000, User, ObjectAceFlags.ObjectAceTypePresent, Guid.Parse("7185fd7e-d6fa-4cf8-b138-ef26da1cb150"),
                Guid.Empty, false, null));
            var source = Descriptor(acl, Other, ControlFlags.DiscretionaryAclAutoInherited);
            var target = Descriptor(new RawAcl(GenericAcl.AclRevision, 0), User, ControlFlags.DiscretionaryAclAutoInherited);
            byte[] sourceBefore = NativeProbeDaclCopy.Bytes(source), targetBefore = NativeProbeDaclCopy.Bytes(target);
            var copied = NativeProbeDaclCopy.Create(sourceBefore, targetBefore);
            Assert.AreEqual(acl.Revision, copied.DiscretionaryAcl.Revision);
            Assert.AreEqual(3, copied.DiscretionaryAcl.Count);
            for (int index = 0; index < 3; index++)
            {
                byte[] expected = NativeProbeDaclCopy.Bytes(acl[index]);
                expected[1] &= unchecked((byte)~(byte)AceFlags.Inherited);
                CollectionAssert.AreEqual(expected, NativeProbeDaclCopy.Bytes(copied.DiscretionaryAcl[index]));
            }
            Assert.AreEqual(unchecked((int)0x80000000U), ((KnownAce)copied.DiscretionaryAcl[0]).AccessMask);
            Assert.AreEqual(0x10000000, ((KnownAce)copied.DiscretionaryAcl[1]).AccessMask);
            // This exact buffer goes to SetFileSecurityW with DACL-only information;
            // avoid FileSystemAccessRule/DirectorySecurity rewriting native ACE flags.
            var roundTrip = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(copied), 0);
            CollectionAssert.AreEqual(NativeProbeDaclCopy.Bytes(copied.DiscretionaryAcl), NativeProbeDaclCopy.Bytes(roundTrip.DiscretionaryAcl));
            Assert.AreEqual(ControlFlags.DiscretionaryAclProtected, roundTrip.ControlFlags & ControlFlags.DiscretionaryAclProtected);
            Assert.AreEqual(User, copied.Owner);
            Assert.AreEqual(target.Group, copied.Group);
            Assert.AreEqual(ControlFlags.DiscretionaryAclProtected,
                copied.ControlFlags & (ControlFlags.DiscretionaryAclProtected | ControlFlags.DiscretionaryAclAutoInherited));
            CollectionAssert.AreEqual(sourceBefore, NativeProbeDaclCopy.Bytes(source));
            CollectionAssert.AreEqual(targetBefore, NativeProbeDaclCopy.Bytes(target));
        }

        [TestMethod]
        public void ReadbackAllowsOnlyObservedAdditionOfAutoInheritedWithoutChangingDacl()
        {
            var expected = CopyFullControl();
            var actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            NativeProbeDaclCopy.VerifyReadback(expected, actual);
            actual.SetFlags(actual.ControlFlags | ControlFlags.DiscretionaryAclAutoInherited);
            NativeProbeDaclCopy.VerifyReadback(expected, actual);
            CollectionAssert.AreEqual(NativeProbeDaclCopy.Bytes(expected.DiscretionaryAcl), NativeProbeDaclCopy.Bytes(actual.DiscretionaryAcl));
            Assert.AreEqual(ControlFlags.DiscretionaryAclAutoInherited, expected.ControlFlags ^ actual.ControlFlags);
            actual.SetFlags(actual.ControlFlags & ~ControlFlags.DiscretionaryAclProtected);
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
        }

        [TestMethod]
        public void ReadbackRejectsChangedMaskOwnerOrderAndInheritedFlag()
        {
            var expected = CopyFullControl();
            var actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            ((KnownAce)actual.DiscretionaryAcl[0]).AccessMask = (int)FileSystemRights.Read;
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
            actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            actual.Owner = Other;
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
            actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            actual.Group = User;
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
            actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            actual.DiscretionaryAcl[0].AceFlags |= AceFlags.Inherited;
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
            actual = new RawSecurityDescriptor(NativeProbeDaclCopy.Bytes(expected), 0);
            var first = actual.DiscretionaryAcl[0]; actual.DiscretionaryAcl.RemoveAce(0); actual.DiscretionaryAcl.InsertAce(1, first);
            Assert.ThrowsException<InvalidOperationException>(() => NativeProbeDaclCopy.VerifyReadback(expected, actual));
        }

        [TestMethod]
        public void SourceGrantMustBeDirectForCurrentSidAndMustNeverBeInvented()
        {
            var acl = new RawAcl(GenericAcl.AclRevision, 1);
            var ace = new CommonAce(AceFlags.ContainerInherit | AceFlags.InheritOnly, AceQualifier.AccessAllowed,
                0x10000000, User, false, null);
            acl.InsertAce(0, ace);
            var source = Descriptor(acl, Other);
            Assert.IsFalse(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(source), User));
            ace.AceFlags &= ~AceFlags.InheritOnly;
            Assert.IsTrue(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(source), User));
            Assert.IsFalse(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(source), Other));
            ace.AccessMask = unchecked((int)0x80000000U);
            Assert.IsFalse(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(source), User));
            ace.AccessMask = (int)FileSystemRights.FullControl;
            Assert.IsTrue(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(source), User));
            var denied = new RawAcl(GenericAcl.AclRevision, 1);
            denied.InsertAce(0, new CommonAce(AceFlags.None, AceQualifier.AccessDenied, 0x10000000, User, false, null));
            Assert.IsFalse(NativeProbeDaclCopy.SourceAllowsDirectFullControl(NativeProbeDaclCopy.Bytes(Descriptor(denied, User)), User));
        }

        [TestMethod]
        public void NullDaclIsRejectedAndTargetSystemAclIsPreservedWithoutMutation()
        {
            var noDacl = new RawSecurityDescriptor(ControlFlags.None, Other, Other, null, null);
            var target = Descriptor(new RawAcl(GenericAcl.AclRevision, 0), User);
            Assert.ThrowsException<ArgumentException>(() => NativeProbeDaclCopy.Create(NativeProbeDaclCopy.Bytes(noDacl), NativeProbeDaclCopy.Bytes(target)));
            var audit = new RawAcl(GenericAcl.AclRevision, 1);
            audit.InsertAce(0, new CommonAce(AceFlags.SuccessfulAccess, AceQualifier.SystemAudit, 1, User, false, null));
            target.SystemAcl = audit;
            target.SetFlags(target.ControlFlags | ControlFlags.SystemAclPresent | ControlFlags.SystemAclProtected);
            byte[] before = NativeProbeDaclCopy.Bytes(target);
            var copied = NativeProbeDaclCopy.Create(NativeProbeDaclCopy.Bytes(CopyFullControl()), before);
            CollectionAssert.AreEqual(NativeProbeDaclCopy.Bytes(audit), NativeProbeDaclCopy.Bytes(copied.SystemAcl));
            Assert.AreEqual(ControlFlags.SystemAclProtected, copied.ControlFlags & ControlFlags.SystemAclProtected);
            CollectionAssert.AreEqual(before, NativeProbeDaclCopy.Bytes(target));
        }

        private static RawSecurityDescriptor CopyFullControl()
        {
            var acl = new RawAcl(GenericAcl.AclRevision, 2);
            acl.InsertAce(0, new CommonAce(AceFlags.Inherited | AceFlags.ContainerInherit, AceQualifier.AccessAllowed,
                (int)FileSystemRights.FullControl, User, false, null));
            acl.InsertAce(1, new CommonAce(AceFlags.None, AceQualifier.AccessAllowed, 0x10000000, Other, false, null));
            return NativeProbeDaclCopy.Create(NativeProbeDaclCopy.Bytes(Descriptor(acl, Other)),
                NativeProbeDaclCopy.Bytes(Descriptor(new RawAcl(GenericAcl.AclRevision, 0), User)));
        }

        private static RawSecurityDescriptor Descriptor(RawAcl acl, SecurityIdentifier owner, ControlFlags extra = ControlFlags.None)
        { return new RawSecurityDescriptor(ControlFlags.DiscretionaryAclPresent | extra, owner, Other, null, acl); }
    }
}
