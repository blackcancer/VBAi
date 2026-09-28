using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie types, dimensions, bornes et refus sans aucun appel de macro.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaProcedureValuesTests
    {
        /// <summary>Copie profonde des vecteurs/matrices JSON avec Null explicite et bornes d'entrée fixes à zéro.</summary>
        [TestMethod]
        public void InputArraysBecomeIndependentZeroBasedVariantSafeArrays()
        {
            var vector = new object[] { 1, "newline\nvalue", true, null };
            var first = new object[] { 1, 2 }; var second = new object[] { 3, 4 };
            object[] captured = VbaProcedureValues.Capture(new object[] { vector, new object[] { first, second }, 4294967295L });
            vector[0] = 99; first[0] = 88;
            Assert.AreEqual(1, ((object[])captured[0])[0]); Assert.AreEqual(DBNull.Value, ((object[])captured[0])[3]);
            Assert.AreEqual(1, ((object[,])captured[1])[0, 0]); Assert.AreEqual(4, ((object[,])captured[1])[1, 1]);
            Assert.AreEqual(0, ((Array)captured[1]).GetLowerBound(0));
            Assert.AreEqual(4294967295m, captured[2]);
            foreach (object bad in new object[] { new Dictionary<string, object>(), new object(), DateTime.UtcNow, double.NaN,
                new object[] { 1, new object[] { 2 } }, new object[] { new object[] { 1 }, new object[] { 1, 2 } },
                new object[] { new object[] { new object[] { 1 } } }, new object[0], new object[] { new object[0] }, new object[1025], new string('x', 16385), Array.CreateInstance(typeof(object), new[] { 2 }, new[] { 1 }) })
                Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Capture(new[] { bad }));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Capture(null));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Capture(new object[31]));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Capture(Enumerable.Repeat<object>(new object[1024], 5).ToArray()));
        }

        /// <summary>La conversion du retour conserve les bornes natives et refuse tout getter ou tableau imbriqué arbitraire.</summary>
        [TestMethod]
        public void ReturnedScalarArraysPreserveNativeBoundsWithoutRetainingComObjects()
        {
            var native = Array.CreateInstance(typeof(object), new[] { 2, 2 }, new[] { 1, -2 });
            native.SetValue(7, 1, -2); native.SetValue("text", 1, -1); native.SetValue(true, 2, -2); native.SetValue(DBNull.Value, 2, -1);
            var result = VbaProcedureValues.NormalizeReturn(native);
            Assert.AreEqual("Array", result.Kind); CollectionAssert.AreEqual(new[] { 1, -2 }, result.LowerBounds); CollectionAssert.AreEqual(new[] { 2, 2 }, result.Lengths);
            var rows = (object[])result.Value; Assert.AreEqual(7, ((object[])rows[0])[0]); Assert.IsNull(((object[])rows[1])[1]);
            Assert.AreEqual("Empty", VbaProcedureValues.NormalizeReturn(null).Kind); Assert.AreEqual("Null", VbaProcedureValues.NormalizeReturn(DBNull.Value).Kind);
            Assert.AreEqual("Scalar", VbaProcedureValues.NormalizeReturn(7.5m).Kind);
            Assert.AreEqual(0, ((object[])VbaProcedureValues.NormalizeReturn(new int[0]).Value).Length);
            foreach (object bad in new object[] { new object(), DateTime.UtcNow, double.PositiveInfinity, new object[] { new object[] { 1 } }, new object[4097], new int[1, 1, 1] })
                Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureValues.NormalizeReturn(bad));
        }

        /// <summary>Les noms sont ordonnés pour Excel.Run et les Optional omis sont Missing, sans passage natif de noms.</summary>
        [TestMethod]
        public void NamedOptionalBindingUsesExactByValTypesAndRejectsCoercionLoss()
        {
            const string signature = "Public Function F(ByVal first As Long, Optional ByVal second As Variant = 0, Optional ByVal third As String = \"ok\") As Variant\nEnd Function";
            var values = VbaProcedureValues.Capture(new object[] { "value", 7 });
            var bound = VbaProcedureValues.Bind(signature, 1, "F", values, new[] { "THIRD", "first" });
            Assert.AreEqual(7, bound[0]); Assert.AreSame(Type.Missing, bound[1]); Assert.AreEqual("value", bound[2]);
            foreach (string declaration in new[] { "ByRef first As Long", "first As Long", "ByVal first() As Long", "ByVal first As Object", "ParamArray first() As Variant" })
                Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureValues.Bind("Public Sub F(" + declaration + ")\nEnd Sub", 1, "F", new object[] { 7 }, null));
            foreach (object value in new object[] { 7.5, 2147483648m, DBNull.Value, true, "7", new object[] { 1 } })
                Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(signature, 1, "F", new[] { value }, null));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(signature, 1, "F", new object[0], null));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(signature, 1, "F", new object[] { 7 }, new[] { "absent" }));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(signature, 1, "F", new object[] { 7, 8 }, new[] { "first", "FIRST" }));
            Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureValues.Bind("#Const X=True\n" + signature, 2, "F", new object[] { 7 }, null));
        }
    }
}
