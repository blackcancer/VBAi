using System;
using System.Globalization;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbaProcedureValuesBoundaryTests
    {
        private static object BindScalar(string type, object value) => VbaProcedureValues.Bind("Public Sub F(ByVal x As " + type + ")\nEnd Sub", 1, "F", new[] { value }, null)[0];
    }
    /// <summary>One explicit scalar binding and its expected native type.</summary>
    internal sealed class ScalarCase
    {
        internal readonly string Type;
        internal readonly object Value, Expected;
        internal ScalarCase(string type, object value, object expected) { Type = type; Value = value; Expected = expected; }
    }
    /// <summary>Exercises the conversion exception boundary, including errors outside the two translated numeric failures.</summary>
    internal sealed class ConversionFailureValue : IConvertible
    {
        private readonly Exception error;
        internal ConversionFailureValue(Exception error) { this.error = error; }
        public TypeCode GetTypeCode() => TypeCode.Decimal;
        public decimal ToDecimal(IFormatProvider provider) { throw error; }
        public bool ToBoolean(IFormatProvider provider) => throw new NotSupportedException();
        public byte ToByte(IFormatProvider provider) => throw new NotSupportedException();
        public char ToChar(IFormatProvider provider) => throw new NotSupportedException();
        public DateTime ToDateTime(IFormatProvider provider) => throw new NotSupportedException();
        public double ToDouble(IFormatProvider provider) => throw new NotSupportedException();
        public short ToInt16(IFormatProvider provider) => throw new NotSupportedException();
        public int ToInt32(IFormatProvider provider) => throw new NotSupportedException();
        public long ToInt64(IFormatProvider provider) => throw new NotSupportedException();
        public sbyte ToSByte(IFormatProvider provider) => throw new NotSupportedException();
        public float ToSingle(IFormatProvider provider) => throw new NotSupportedException();
        public string ToString(IFormatProvider provider) => throw new NotSupportedException();
        public object ToType(Type conversionType, IFormatProvider provider) => throw new NotSupportedException();
        public ushort ToUInt16(IFormatProvider provider) => throw new NotSupportedException();
        public uint ToUInt32(IFormatProvider provider) => throw new NotSupportedException();
        public ulong ToUInt64(IFormatProvider provider) => throw new NotSupportedException();
    }
}