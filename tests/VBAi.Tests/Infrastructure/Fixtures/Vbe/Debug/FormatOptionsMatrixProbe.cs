namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;

    /// <summary>Simule le catalogue global de catégories et leurs sélections sans remplacer l'orchestration Options.</summary>
    internal sealed class FormatOptionsMatrixProbe : VbeDebugWindows.IWritableOptionsProbe, VbeDebugWindows.IFormatCategoriesOptionsProbe
    {
        internal readonly WritableOptionsMatrixProbe Inner = new WritableOptionsMatrixProbe();
        internal readonly List<VbeDebugWindows.OptionsFormatCategory> Categories = new List<VbeDebugWindows.OptionsFormatCategory>();
        internal int Selections;
        internal FormatOptionsMatrixProbe()
        {
            Inner.Names[0] = "Editor Format";
            Inner.Items.Clear();
            Inner.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Code Colors", Type = "ControlType.List", Value = "Normal", Choices = new[] { "Normal", "Comment" } });
            var palette = new VbeDebugWindows.OptionsControl { Name = "Foreground", Type = "ControlType.ComboBox" };
            VbeDebugWindows.DescribeOptionsNativeChoices(palette, new[] { "Automatic", "", "" }, 0, "Automatic");
            Inner.Items.Add(palette);
            foreach (string category in new[] { "Normal", "Comment" })
                Categories.Add(new VbeDebugWindows.OptionsFormatCategory { Category = category, Palettes = new[] { new VbeDebugWindows.OptionsControl { Name = "Foreground", Value = "Automatic" } } });
        }
        public IntPtr Dialog() => Inner.Dialog();
        public IList<string> Tabs(IntPtr dialog) => Inner.Tabs(dialog);
        public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int tabIndex) => Inner.Controls(dialog, tabIndex);
        public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => Inner.ErrorChoices(dialog);
        public void Close(IntPtr dialog) => Inner.Close(dialog);
        public void Pause(int milliseconds) => Inner.Pause(milliseconds);
        public void Write(IntPtr dialog, int tabIndex, string name, string type, object value) => Inner.Write(dialog, tabIndex, name, type, value);
        public void Accept(IntPtr dialog) => Inner.Accept(dialog);
        public IList<VbeDebugWindows.OptionsFormatCategory> FormatCategories(IntPtr dialog, int tabIndex) => Categories;
        public void SelectFormatCategory(IntPtr dialog, int tabIndex, string category)
        {
            if (Categories.Count(x => x.Category == category) != 1) throw new InvalidOperationException("Unknown category");
            Selections++; Inner.Items[0].Value = category;
        }
        internal Request Request()
        {
            var request = Inner.Request(); request.Property = "Foreground"; request.Value = "NativeIndex:1"; request.Query = "Comment";
            dynamic read = VbeDebugWindows.ReadVbeOptions(this); Inner.Open = true; Inner.Closes = 0; request.ExpectedOptionsVersion = read.OptionsVersion;
            return request;
        }
    }
}
