namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Source;
            public VbeForms Service;
            public Request Request(string path = null, string name = "Copy")
            {
                return new Request
                {
                    Project = Project.Name,
                    Form = Form.Name,
                    ControlPath = path ?? "Controls/" + Source.Name,
                    NewName = name,
                    ExpectedTreeVersion = ((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion
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
            var source = form.Designer.Controls.AddExisting(type, type + "1");
            source.Caption = "Source";
            source.Left = 3;
            source.Top = 4;
            source.Width = 50;
            source.Height = 16;
            return new Fixture
            {
                Project = project,
                Form = form,
                Source = source,
                Service = new VbeForms(vbe)
            };
        }
    }
}
