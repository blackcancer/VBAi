using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Reviews only the empty default template in the owned /a process; never reads code text or template files.</summary>
        internal static object ReviewEmptyStartupNormalTemplate(object template, object normal, Action<object> release = null)
        {
            Assert.IsTrue(VbeProjectHostPath.SameProject(template, normal), "An arbitrary foreign template is not the startup Normal object.");
            using (var references = new VbaTestComReferences(release))
            {
                object project = references.Add((object)((dynamic)template).VBProject);
                Assert.IsNotNull(project);
                string name = Convert.ToString(((dynamic)project).Name);
                int mode = Convert.ToInt32(((dynamic)project).Mode), protection = Convert.ToInt32(((dynamic)project).Protection);
                Assert.AreEqual("Normal", name, true);
                Assert.AreEqual(2, mode, "Normal must be in design mode.");
                Assert.AreEqual(0, protection, "A protected Normal cannot be reviewed.");
                var components = references.Items((object)((dynamic)project).VBComponents).Take(2).ToArray();
                Assert.AreEqual(1, components.Length, "Normal must contain only its empty document class.");
                Assert.AreEqual(100, Convert.ToInt32(((dynamic)components[0]).Type));
                Assert.AreEqual("ThisDocument", Convert.ToString(((dynamic)components[0]).Name), true);
                object module = references.Add((object)((dynamic)components[0]).CodeModule);
                Assert.IsNotNull(module);
                int lines = Convert.ToInt32(((dynamic)module).CountOfLines);
                Assert.AreEqual(0, lines, "Normal with any VBA code is refused before changing suppression.");
                return new { TemplatePath = Convert.ToString(((dynamic)template).FullName), ProjectName = name,
                    Mode = mode, Protection = protection, ComponentCount = components.Length, CodeLineCount = lines,
                    CodeTextRead = false, TemplateFileRead = false, EmptyStartupNormal = true };
            }
        }
    }
}
