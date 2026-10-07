namespace VBAi.Tests.Unit
{
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using VBAi;

    public sealed partial class ProviderProtocolTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static IDictionary<string, object> Obj(object value)
        {
            return ClaudeProtocol.Object(value);
        }
    }
}
