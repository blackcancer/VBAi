namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Forms;

    // In-memory IDataObject: exercises the same clipboard serialization contract without opening the clipboard.
    internal sealed class DesignerDataFixture : IDataObject
    {
        internal readonly Dictionary<string,object> Values = new Dictionary<string,object>();
        public object GetData(string format, bool autoConvert) { return Values.TryGetValue(format,out object value)?value:null; }
        public object GetData(string format) { return GetData(format,false); }
        public object GetData(Type format) { return GetData(format.FullName,false); }
        public bool GetDataPresent(string format, bool autoConvert) { return Values.ContainsKey(format); }
        public bool GetDataPresent(string format) { return GetDataPresent(format,false); }
        public bool GetDataPresent(Type format) { return GetDataPresent(format.FullName,false); }
        public string[] GetFormats(bool autoConvert) { return Values.Keys.ToArray(); }
        public string[] GetFormats() { return GetFormats(false); }
        public void SetData(string format,bool autoConvert,object data) { Values[format]=data; }
        public void SetData(string format,object data) { Values[format]=data; }
        public void SetData(Type format,object data) { Values[format.FullName]=data; }
        public void SetData(object data) { Values[data.GetType().FullName]=data; }
    }
}
