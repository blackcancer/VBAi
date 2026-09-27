namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Web.Script.Serialization;
    using System.Threading.Tasks;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class LlmVbeAsyncValidationTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static async Task Failure(LlmVbeTools tools, string name, string arguments, string fragment)
        {
            string serialized = await tools.InvokeAsync(name, arguments);
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok, name + " accepted invalid arguments");
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
