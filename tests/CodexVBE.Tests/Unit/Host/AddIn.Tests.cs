namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie l’arrêt de l’add-in et la propriété des fenêtres natives.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HostSettingsCoverageTests
    {
        /// <summary>Ferme la fenêtre native une fois et rend l’arrêt répétable.</summary>
        [TestMethod]
        [STATestMethod]
        public void AddInShutdownClosesNativeWindowAndCanRunTwice()
        {
            var addIn = new AddIn();
            var native = new FakeNativeWindow();
            var dispatcher = new Control();
            Set(addIn, "nativeChatWindow", native);
            Set(addIn, "dispatcher", dispatcher);
            object[] custom = null;
            addIn.OnBeginShutdown(ref custom);
            Assert.AreEqual(1, native.CloseCount);
            Assert.IsTrue(dispatcher.IsDisposed);
            Assert.IsNull(Field<object>(addIn, "nativeChatWindow"));
            Assert.IsNull(Field<object>(addIn, "dispatcher"));
            addIn.OnDisconnection(0, ref custom);
            Assert.AreEqual(1, native.CloseCount);
        }

        /// <summary>Utilise la fenêtre principale du VBE comme propriétaire des dialogues.</summary>
        [TestMethod]
        public void AddInUsesTheVbeMainWindowAsDialogOwner()
        {
            var addIn = new AddIn();
            Set(addIn, "vbe", new FakeOwnerHost { MainWindow = new FakeMainWindow { HWnd = 12345 } });
            var owner = (IWin32Window)Call(addIn, "VbeOwner");
            Assert.AreEqual(new IntPtr(12345), owner.Handle);
            object[] custom = null;
            addIn.OnDisconnection(0, ref custom);
            Assert.IsNull(Field<object>(addIn, "vbe"));
        }

        /// <summary>Libère les ressources de l’add-in même si la fenêtre native refuse la fermeture.</summary>
        [TestMethod]
        public void AddInShutdownContinuesWhenTheNativeWindowRejectsClose()
        {
            var addIn = new AddIn();
            var dispatcher = new Control();
            Set(addIn, "dispatcher", dispatcher);
            Set(addIn, "nativeChatWindow", new RejectingNativeWindow());
            Set(addIn, "addIn", new object ());
            object[] custom = null;
            addIn.OnBeginShutdown(ref custom);
            Assert.IsTrue(dispatcher.IsDisposed);
            Assert.IsNull(Field<object>(addIn, "nativeChatWindow"));
            Assert.IsNull(Field<object>(addIn, "addIn"));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie les métadonnées COM et les callbacks sans connexion active.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HostSettingsWindowTests
    {
        /// <summary>Conserve des callbacks inoffensifs lorsque l’objet AddIn n’a pas été initialisé par le VBE.</summary>
        [TestMethod]
        public void AddInMetadataAndNoOpLifecycleCallbacksRemainSafeWithoutConnection()
        {
            Assert.AreEqual("CodexVBE.AddIn", ((ProgIdAttribute)Attribute.GetCustomAttribute(typeof(AddIn), typeof(ProgIdAttribute))).Value);
            var addin = (AddIn)FormatterServices.GetUninitializedObject(typeof(AddIn));
            object[] custom = new object[0];
            addin.OnAddInsUpdate(ref custom);
            addin.OnStartupComplete(ref custom);
            addin.OnBeginShutdown(ref custom);
        }
    }
}
