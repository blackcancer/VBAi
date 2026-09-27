namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class WinFormsDesignerTests
    {
        private static T Find<T>(Control root, string name)
            where T : Control
        {
            var matches = root.Controls.Find(name, true);
            Assert.AreEqual(1, matches.Length, "Missing or duplicated designer control: " + name);
            return (T)matches[0];
        }
    }
}
