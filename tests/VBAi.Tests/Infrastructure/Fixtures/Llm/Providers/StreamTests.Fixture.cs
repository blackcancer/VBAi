namespace VBAi.Tests.Unit
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Web.Script.Serialization;

    public sealed partial class StreamTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static IDictionary<string, object> Obj(object value)
        {
            return (IDictionary<string, object>)value;
        }

        private static MemoryStream Events(params object[] events)
        {
            var text = string.Join("", events.Select(e => "data: " + (e is string ? e : Json.Serialize(e)) + "\n\n"));
            return new MemoryStream(Encoding.UTF8.GetBytes(text));
        }
    }
}
