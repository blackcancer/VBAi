namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using CodexVBE;

    public sealed partial class VbeFormsFitContentTests
    {
        private static FitFixture Create(bool nested = false)
        {
            var f = new FitFixture();
            f.Project = new VbeFormsTests.FakeProject();
            f.Form = new VbeFormsCoverageTests.Form();
            f.Project.VBComponents.Add(f.Form);
            var host = new VbeFormsTests.FakeVbe(); host.VBProjects.Add(f.Project);
            f.Service = new VbeForms(host);
            f.Target = nested ? new VbeFormsCoverageTests.Node { Name = "Frame1", ClassName = "Frame", Parent = f.Form.Designer } : f.Form.Designer;
            f.Children = new[] { new VbeFormsCoverageTests.Node { Name = "Label1", ClassName = "Label", Parent = f.Target } };
            f.Target.Controls = f.Children;
            var child = f.Children[0];
            child.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] {
                new VbeFormsCoverageTests.LiveProperty("Left", typeof(double), () => f.ChildLeft),
                new VbeFormsCoverageTests.LiveProperty("Top", typeof(double), () => f.ChildTop),
                new VbeFormsCoverageTests.LiveProperty("Width", typeof(double), () => f.ChildWidth),
                new VbeFormsCoverageTests.LiveProperty("Height", typeof(double), () => f.ChildHeight) });
            f.Target.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] {
                new VbeFormsCoverageTests.LiveProperty("Controls", typeof(object), () => f.Target.Controls),
                new VbeFormsCoverageTests.LiveProperty("Width", typeof(double), () => f.Width, value => f.Write("Width", value)),
                new VbeFormsCoverageTests.LiveProperty("Height", typeof(double), () => f.Height, value => f.Write("Height", value)),
                new VbeFormsCoverageTests.LiveProperty("InsideWidth", typeof(double), () => f.Width - f.BorderX),
                new VbeFormsCoverageTests.LiveProperty("InsideHeight", typeof(double), () => f.Height - f.BorderY),
                new VbeFormsCoverageTests.LiveProperty("ScrollWidth", typeof(double), () => f.ScrollWidth, value => f.Write("ScrollWidth", value)),
                new VbeFormsCoverageTests.LiveProperty("ScrollHeight", typeof(double), () => f.ScrollHeight, value => f.Write("ScrollHeight", value)) });
            if (nested)
            {
                f.Form.Designer.Controls = new[] { f.Target };
                f.Form.Designer.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] {
                    new VbeFormsCoverageTests.LiveProperty("Controls", typeof(object), () => f.Form.Designer.Controls) });
                f.Path = "Controls/Frame1";
            }
            return f;
        }

        private sealed class FitFixture
        {
            internal VbeForms Service;
            internal VbeFormsTests.FakeProject Project;
            internal VbeFormsCoverageTests.Form Form;
            internal VbeFormsCoverageTests.Node Target;
            internal VbeFormsCoverageTests.Node[] Children;
            internal string Path;
            internal double Width = 300, Height = 200, BorderX = 10, BorderY = 30;
            internal double ChildLeft = 20, ChildTop = 10, ChildWidth = 50, ChildHeight = 25;
            internal double ScrollWidth = 500, ScrollHeight = 400;
            internal int Writes;
            internal bool Ignore, Fail, ChangeChild, ChangeDecoration;
            internal Request Request(string action = "fit_container")
            {
                return new Request { Project = Project.Name, Form = Form.Name, ControlPath = Path,
                    Action = action, Left = 5, Top = 7,
                    ExpectedTreeVersion = (string)((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion };
            }
            internal void Write(string property, object value)
            {
                Writes++;
                if (Fail) throw new InvalidOperationException("Native setter failed");
                if (Ignore) return;
                double number = Convert.ToDouble(value);
                if (property == "Width") Width = number;
                if (property == "Height") Height = number;
                if (property == "ScrollWidth") ScrollWidth = number;
                if (property == "ScrollHeight") ScrollHeight = number;
                if (ChangeChild) ChildLeft++;
                if (ChangeDecoration) BorderX++;
            }
        }
    }
}
