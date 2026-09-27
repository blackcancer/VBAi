namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeChatReferenceTests
    {
        private static VbeSession Session()
        {
            var vbe = new VbeSessionTests.FakeVbe();
            var project = new VbeSessionTests.FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\ReferenceTest.xlsm",
                Mode = 2
            };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = new VbeSessionTests.FakeModule("Sub Example()\r\nEnd Sub") });
            vbe.VBProjects.Add(project);
            return new VbeSession(vbe);
        }
    }
}
