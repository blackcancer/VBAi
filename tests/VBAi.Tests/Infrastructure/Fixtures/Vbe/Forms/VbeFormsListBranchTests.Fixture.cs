using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Reflection;
namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsTests
    {
        private static void WithList(string type, Action<Fixture, FakeControl, Request> run, int rows = 2)
        {
            var f = NewFixture();
            var c = f.Form.Designer.Controls.AddExisting("Choices");
            for (int i = 0; i < rows; i++) c.ListRows.Add(new object[] { "item" + i });
            var provider = new NamedControlProvider(type);
            TypeDescriptor.AddProvider(provider, c);
            try
            {
                dynamic before = f.Service.ListItems(new Request { Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/Choices", Limit = 64 });
                var r = new Request
                {
                    Project = f.Project.Name,
                    Form = f.Form.Name,
                    ControlPath = "Controls/Choices",
                    ExpectedTreeVersion = before.TreeVersion,
                    ExpectedListVersion = before.ListVersion,
                    Text = "new",
                    RowIndex = 0,
                    Limit = 64
                };
                run(f, c, r);
            }
            finally { TypeDescriptor.RemoveProvider(provider, c); }
        }

        private static Request RequiredListRequest()
        {
            return new Request { Project = "P", Form = "F", ControlPath = "Controls/C", ExpectedTreeVersion = "tree", ExpectedListVersion = "list", Text = "", RowIndex = 0, ExpectedSha256 = "sha", Items = new string[0] };
        }

        private static void MissingFields(Action<Request> action, params string[] fields)
        {
            foreach (var field in fields)
            {
                var r = RequiredListRequest();
                typeof(Request).GetProperty(field).SetValue(r, null);
                Assert.ThrowsException<ArgumentException>(() => action(r), field);
            }
        }

        public sealed class CellTextFailure
        {
            private readonly bool invocation;
            public CellTextFailure(bool invocation) { this.invocation = invocation; }
            public override string ToString()
            {
                if (invocation) throw new TargetInvocationException(null);
                throw new InvalidOperationException("Scalar conversion failed");
            }
        }
    }
}