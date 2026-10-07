namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Reflection;

    /// <summary>Fournit des objets hôte simulés et des accès réfléchis aux tests de configuration.</summary>
    public sealed partial class HostSettingsCoverageTests
    {
        /// <summary>Options de réflexion ciblant les champs d’instance privés.</summary>
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
        /// <summary>Lit un champ privé d’un objet simulé ou d’une fenêtre.</summary>
        /// <typeparam name="T">Type attendu du champ.</typeparam>
        /// <param name="instance">Objet qui porte le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        private static T Field<T>(object instance, string name)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(instance);
        }

        /// <summary>Modifie un champ privé utilisé comme point d’injection dans un test.</summary>
        /// <param name="instance">Objet qui porte le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <param name="value">Valeur à affecter.</param>
        private static void Set(object instance, string name, object value)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            field.SetValue(instance, value);
        }

        /// <summary>Appelle une méthode privée par réflexion.</summary>
        /// <param name="instance">Objet cible de l’appel.</param>
        /// <param name="name">Nom de la méthode.</param>
        /// <param name="arguments">Arguments à transmettre.</param>
        /// <returns>Valeur renvoyée par la méthode.</returns>
        private static object Call(object instance, string name, params object[] arguments)
        {
            var method = instance.GetType().GetMethod(name, PrivateInstance);
            Assert.IsNotNull(method, name);
            return method.Invoke(instance, arguments);
        }

        /// <summary>Expose les barres de commandes de l’hôte simulé.</summary>
        public sealed class FakeHost
        {
            /// <summary>Barres de commandes renvoyées à l’inspection.</summary>
            /// <value>Objets représentant les barres de l’hôte.</value>
            public object[] CommandBars { get; set; }
        }

        /// <summary>Barre simulée avec son type et ses contrôles.</summary>
        public sealed class FakeBar
        {
            /// <summary>Type numérique de la barre.</summary>
            /// <value>Valeur lue par le code de détection des menus.</value>
            public int Type { get; set; }
            /// <summary>Contrôles contenus dans la barre.</summary>
            /// <value>Objets de contrôle simulés.</value>
            public object[] Controls { get; set; }
        }

        /// <summary>Barre dont la lecture du type échoue pour tester la tolérance aux exceptions COM.</summary>
        public sealed class InvalidBar
        {
            /// <summary>Provoque une exception lorsque le type de la barre est lu.</summary>
            /// <value>La lecture lève toujours <see cref="InvalidOperationException"/>.</value>
            public int Type
            {
                get
                {
                    throw new InvalidOperationException("no type");
                }
            }
        }

        /// <summary>Contrôle de menu simulé exposant son libellé.</summary>
        public sealed class FakeControl
        {
            /// <summary>Texte visible du contrôle.</summary>
            /// <value>Libellé lu pour reconnaître les menus.</value>
            public string Caption { get; set; }
        }

        /// <summary>Contrôle dont la lecture du libellé échoue pour couvrir les erreurs d’interop.</summary>
        public sealed class InvalidControl
        {
            /// <summary>Provoque une exception lorsque le libellé est lu.</summary>
            /// <value>La lecture lève toujours <see cref="InvalidOperationException"/>.</value>
            public string Caption
            {
                get
                {
                    throw new InvalidOperationException("no caption");
                }
            }
        }

        /// <summary>Fenêtre native de test qui compte les demandes de fermeture.</summary>
        public sealed class FakeNativeWindow
        {
            /// <summary>Nombre d’appels à <see cref="Close"/>.</summary>
            /// <value>Compteur incrémenté à chaque fermeture.</value>
            public int CloseCount { get; private set; }

            /// <summary>Enregistre une demande de fermeture.</summary>
            public void Close()
            {
                CloseCount++;
            }
        }

        /// <summary>Hôte simulé qui expose une fenêtre principale.</summary>
        public sealed class FakeOwnerHost
        {
            /// <summary>Fenêtre principale fournie au code d’attachement du dialogue.</summary>
            /// <value>Fenêtre simulée de l’hôte.</value>
            public FakeMainWindow MainWindow { get; set; }
        }

        /// <summary>Fenêtre principale simulée exposant son handle natif.</summary>
        public sealed class FakeMainWindow
        {
            /// <summary>Handle Win32 de la fenêtre hôte.</summary>
            /// <value>Valeur entière du handle.</value>
            public long HWnd { get; set; }
        }

        /// <summary>Fenêtre dont la fermeture échoue pour simuler un hôte déjà arrêté.</summary>
        public sealed class RejectingNativeWindow
        {
            /// <summary>Lève une exception à la demande de fermeture.</summary>
            public void Close()
            {
                throw new InvalidOperationException("VBE closed first");
            }
        }
    }
}
