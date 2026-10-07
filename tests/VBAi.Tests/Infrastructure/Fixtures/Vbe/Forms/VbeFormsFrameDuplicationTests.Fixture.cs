namespace VBAi.Tests.Unit
{
    using VBAi;

    public sealed partial class VbeFormsFrameDuplicationTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Frame;
            public VbeForms Service;
            public Request Request(string newName = "FrameCopy")
            {
                return new Request
                {
                    Project = Project.Name,
                    Form = Form.Name,
                    ControlPath = "Controls/Frame1",
                    NewName = newName,
                    ExpectedTreeVersion = ((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion
                };
            }
        }

        private static Fixture Create()
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var frame = form.Designer.Controls.AddExisting("Frame", "Frame1");
            frame.Caption = "Group";
            frame.Left = 4;
            frame.Top = 6;
            frame.Width = 90;
            frame.Height = 60;
            return new Fixture
            {
                Project = project,
                Form = form,
                Frame = frame,
                Service = new VbeForms(vbe)
            };
        }
    }
}
