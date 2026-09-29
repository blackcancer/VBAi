using System;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Liaison positionnelle complète et refus des signatures ParamArray invalides.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaProcedureParamArrayTests
    {
        /// <summary>Conserve un argument Excel.Run par valeur, y compris les tableaux et Null imbriqués.</summary>
        [TestMethod]
        public void VariadicValuesRemainDistinctAfterFixedByValCoercion()
        {
            foreach (string tail in new[] { "ParamArray values()", "ParamArray values() As Variant" })
            {
                string source = "Public Function F(ByVal first As Integer, " + tail + ") As Variant\nEnd Function";
                var captured = VbaProcedureValues.Capture(new object[] { 7, "quoted\"", null, true,
                    new object[] { 1, 2 }, new object[] { new object[] { 3, 4 } } });
                var bound = VbaProcedureValues.Bind(source, 1, "F", captured, null);
                Assert.AreEqual(6, bound.Length); Assert.AreEqual((short)7, bound[0]);
                Assert.AreEqual("quoted\"", bound[1]); Assert.AreEqual(DBNull.Value, bound[2]);
                Assert.AreSame(captured[4], bound[4]); Assert.AreSame(captured[5], bound[5]);
                Assert.AreEqual(0, ((Array)bound[5]).GetLowerBound(0));
            }
        }

        /// <summary>Le ParamArray vide n'ajoute pas d'argument Missing ou de tableau synthétique.</summary>
        [TestMethod]
        public void EmptyAndThirtyArgumentCallsKeepTheExactTransportArity()
        {
            const string source = "Public Function F(ParamArray values() As Variant) As Variant\nEnd Function";
            Assert.AreEqual(0, VbaProcedureValues.Bind(source, 1, "F", new object[0], null).Length);
            Assert.AreEqual(0, VbaProcedureValues.Bind(source, 1, "F", new object[0], new string[0]).Length);
            var thirty = Enumerable.Range(1, 30).Cast<object>().ToArray();
            CollectionAssert.AreEqual(thirty, VbaProcedureValues.Bind(source, 1, "F", thirty, null));
            string prefix = string.Join(", ", Enumerable.Range(1, 30).Select(i => "ByVal arg" + i + " As Long"));
            CollectionAssert.AreEqual(thirty, VbaProcedureValues.Bind("Public Sub F(" + prefix + ", ParamArray tail())\nEnd Sub", 1, "F", thirty, null));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[31], null));
        }

        /// <summary>Refuse toutes les variantes interdites par VBA avant la livraison native.</summary>
        [TestMethod]
        public void InvalidVariadicSignaturesAreRejectedBeforeBinding()
        {
            foreach (string parameters in ParamArrayMatrix.RejectedSignatures())
                Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureValues.Bind(
                    "Public Sub F(" + parameters + ")\nEnd Sub", 1, "F", new object[0], null), parameters);
        }

        /// <summary>Un préfixe reste obligatoire et typé, et aucun argument nommé n'est accepté.</summary>
        [TestMethod]
        public void FixedPrefixesAndPositionalOnlyCallsAreEnforced()
        {
            const string source = "Public Function F(ByVal first As Byte, ParamArray values()) As Variant\nEnd Function";
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[0], null));
            foreach (object invalid in new object[] { 256, -1, "7", true, DBNull.Value, new object[] { 1 } })
                Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new[] { invalid }, null));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[] { 7 }, new[] { "first" }));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureValues.Bind(source, 1, "F", new object[] { 7 }, new[] { "values" }));
        }
    }
}
