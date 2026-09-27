namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class ProviderProtocolTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static IDictionary<string, object> Obj(object value)
        {
            return ClaudeProtocol.Object(value);
        }
    }
}
