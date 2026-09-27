namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class LlmVbeToolContractTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer
        {
            MaxJsonLength = 10 * 1024 * 1024
        };
        private static IDictionary<string, object> Dict(object value)
        {
            return (IDictionary<string, object>)value;
        }

        private static void IsFailure(string serialized, string fragment)
        {
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
