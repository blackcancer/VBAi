namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsCoreBranchTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Control;
            public VbeForms Service;
            public string Version => (string)((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion;

            public Request Request(string path)
            {
                return new Request
                {
                    Project = Project.Name,
                    Form = Form.Name,
                    ControlPath = path,
                    ExpectedTreeVersion = Version
                };
            }
        }

        private static Fixture Create(string type)
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture
            {
                Project = project,
                Form = form,
                Control = form.Designer.Controls.AddExisting(type, type + "1"),
                Service = new VbeForms(vbe)
            };
        }
    }
}
