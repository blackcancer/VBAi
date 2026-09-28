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
            foreach (string declaration in new[] { "ByRef first As Long", "first As Long", "ByVal first() As Long", "ByVal first As Object", "ParamArray first() As Long" })
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

namespace CodexVBE.Tests.Unit
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
    public sealed partial class VbaProcedureValuesBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void ExactLiveSignaturesRejectEveryMalformedHeaderParameterAndReturnSuffix()
        {
            var empty = VbaProcedureValues.Bind("\nStatic Public Sub F()\n\nEnd Sub\n", 2, "F", new object[0], null);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, empty.Length);
            foreach (string source in new[] {
                "DefLng A-Z\nPublic Sub F()\nEnd Sub", "Public\nPublic Sub F()\nEnd Sub",
                "Private Sub F()\nEnd Sub", "Public Function G()\nPublic Sub F()\nEnd Sub",
                "Sub F a\nEnd Sub", "Public Sub F\nEnd Sub", "Public Sub F(ByVal x As Long\nEnd Sub",
                "Public Function F() As\nEnd Function", "Public Function F() Long Variant\nEnd Function",
                "Public Function F() As Object\nEnd Function", "Public Function F() As Long(]\nEnd Function",
                "Public Function F() As Long x y\nEnd Function", "Public Function F() As Long extra\nEnd Function",
                "Public Sub F(ByVal x As Long,)\nEnd Sub", "Public Sub F(Optional)\nEnd Sub",
                "Public Sub F(ByVal)\nEnd Sub", "Public Sub F(ByVal _x As Long)\nEnd Sub",
                "Public Sub F(ByVal x As)\nEnd Sub", "Public Sub F(ByVal x For Long)\nEnd Sub",
                "Public Sub F(ByVal x As Long())\nEnd Sub", "Public Sub F(ByVal x As Long = 1)\nEnd Sub",
                "Public Sub F(ByVal x As Long, ByVal X As Long)\nEnd Sub",
                "Public Sub F()\nEnd Sub\nPrivate Sub F()\nEnd Sub"
            })
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[0], null), source);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => VbaProcedureValues.Bind("\nPublic Sub F()\nEnd Sub", 1, "F", new object[0], null));
            var parameters = string.Join(",", System.Linq.Enumerable.Select(System.Linq.Enumerable.Range(0, 31), i => "ByVal x" + i + " As Variant"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => VbaProcedureValues.Bind("Public Sub F(" + parameters + ")\nEnd Sub", 1, "F", new object[0], null));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, VbaProcedureValues.Bind("Public Function F() As Long()\nEnd Function", 1, "F", new object[0], null).Length);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void FixedArgumentNamesArityAndOptionalDefaultsHaveExactBindingContracts()
        {
            const string source = "Public Sub F(Optional ByVal x As Variant = 0)\nEnd Sub";
            foreach (var values in new object[][] { null, new object[31], new object[] { 1, 2 } })
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", values, null));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[] { 1 }, new string[0]));
            var bound = VbaProcedureValues.Bind(source, 1, "F", new object[0], new string[0]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(System.Type.Missing, bound[0]);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void ExplicitScalarTypesPreserveValuesAndRejectEveryLossyNumericConversion()
        {
            foreach (var pair in new[] {
                new ScalarCase("Variant", null, null), new ScalarCase("String", "text", "text"),
                new ScalarCase("Boolean", true, true), new ScalarCase("Double", 12, 12d),
                new ScalarCase("Single", 12, 12f), new ScalarCase("Byte", 255, (byte)255),
                new ScalarCase("Integer", -32768, (short)-32768), new ScalarCase("Long", 2147483647, 2147483647)
            })
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(pair.Expected, BindScalar(pair.Type, pair.Value));
            var currency = (System.Runtime.InteropServices.CurrencyWrapper)BindScalar("Currency", 922337203685477.5807m);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(922337203685477.5807m, currency.WrappedObject);
            currency = (System.Runtime.InteropServices.CurrencyWrapper)BindScalar("Currency", -922337203685477.5808m);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(-922337203685477.5808m, currency.WrappedObject);
            foreach (var pair in new[] {
                new ScalarCase("Long", null, null), new ScalarCase("Boolean", 1, null),
                new ScalarCase("String", 1, null), new ScalarCase("Single", double.MaxValue, null),
                new ScalarCase("Currency", -922337203685477.5809m, null), new ScalarCase("Currency", 922337203685477.5808m, null),
                new ScalarCase("Currency", 1.00001m, null), new ScalarCase("Byte", -1, null), new ScalarCase("Byte", 256, null),
                new ScalarCase("Integer", -32769, null), new ScalarCase("Integer", 32768, null),
                new ScalarCase("Long", -2147483649m, null), new ScalarCase("Long", 2147483648m, null),
                new ScalarCase("Long", double.MaxValue, null), new ScalarCase("Long", System.DateTime.MinValue, null)
            })
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => BindScalar(pair.Type, pair.Value), pair.Type);
            foreach (System.Exception error in new System.Exception[] { new System.OverflowException("overflow"), new System.InvalidCastException("cast"), new System.FormatException("format") })
            {
                var value = new ConversionFailureValue(error);
                if (error is System.FormatException)
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(error, Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.FormatException>(() => BindScalar("Long", value)));
                else
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(error, Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => BindScalar("Long", value)).InnerException);
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CapturedAndReturnedArraysRejectInvalidRanksBoundsAndNativeObjects()
        {
            foreach (object bad in new object[] {
                new object[] { new int[1, 1] }, new object[] { System.Array.CreateInstance(typeof(object), new[] { 1 }, new[] { 1 }) },
                double.NegativeInfinity, float.PositiveInfinity,
                new object[] { new string('x', 16384), new string('x', 16384), new string('x', 16384), new string('x', 16384), "x" }
            })
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => VbaProcedureValues.Capture(new[] { bad }));
            var finite = VbaProcedureValues.Capture(new object[] { 1.5d, 1.25f });
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1.5d, finite[0]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1.25d, finite[1]);
            var negative = VbaProcedureValues.Capture(new object[] { -2147483649m });
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(-2147483649m, negative[0]);
            var returned = (object[])VbaProcedureValues.NormalizeReturn(new object[] { null, System.DBNull.Value }).Value;
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(returned[0]); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(returned[1]);
            object com = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Scripting.Dictionary", true));
            try
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(System.Runtime.InteropServices.Marshal.IsComObject(com));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => VbaProcedureValues.NormalizeReturn(com));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
        }
    }
}