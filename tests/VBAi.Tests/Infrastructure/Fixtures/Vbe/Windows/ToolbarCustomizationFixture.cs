namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    public sealed partial class ToolbarCustomizationTests
    {
        public sealed class Host { public Bars CommandBars { get; } = new Bars(); }
        public sealed class Bars : List<Bar>
        {
            public bool FailAdd, NullAdd;
            public Action<Bar> Added;
            public Button SourceOverride;
            public Bar Add(string name, int position, bool menuBar, bool temporary)
            { if(FailAdd)throw new InvalidOperationException("native add rejected");if(NullAdd)return null;var bar = new Bar { Owner = this, Name = name, Position = position }; Add(bar);Added?.Invoke(bar); return bar; }
            public Button FindControl(int type, int id) => SourceOverride ?? this.SelectMany(x => x.Controls).FirstOrDefault(x => x.Type == type && x.Id == id);
        }
        public sealed class Bar
        {
            internal Bars Owner;
            public string Name { get; set; }
            public int Type {get;set;}
            public bool BuiltIn { get; set; }
            private bool visible=true;
            public bool FailVisibility, IgnoreDelete;
            public bool Visible { get=>visible; set{if(FailVisibility)throw new InvalidOperationException("visibility rejected");visible=value;} }
            public bool Enabled { get; set; } = true;
            public int Protection { get; set; }
            public int Position { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width => 200;
            public int Height => 20;
            public int RowIndex { get; set; } = 1;
            public Buttons Controls { get; } = new Buttons();
            public void Delete() {if(!IgnoreDelete)Owner.Remove(this);}
        }
        public sealed class Buttons : List<Button>
        {
            public bool FailAdd, FailTag, IgnoreTag;
            public int? AddedId;
            public Action<Button> Copied;
            public new Button this[int oneBased] => base[oneBased - 1];
            public Button Add(int type, int id, object parameter, int before, bool temporary)
            {
                if(FailAdd)throw new InvalidOperationException("native button add rejected");
                var button = new Button { Owner = this, Id = id, Type = type, Temporary = temporary, Caption = "Native command", BuiltIn = true };
                if(AddedId.HasValue)button.Id=AddedId.Value;
                Insert(before - 1, button); return button;
            }
        }
        public sealed class Button
        {
            internal Buttons Owner;
            public int Index => Owner.IndexOf(this) + 1;
            public int Id { get; set; }
            public int Type { get; set; }
            public string Caption { get; set; }
            private string tag="";
            public string Tag { get=>tag;set{if(Owner.FailTag)throw new InvalidOperationException("native tag rejected");if(!Owner.IgnoreTag)tag=value;} }
            public bool FailDelete, IgnoreDelete;
            public bool BuiltIn { get; set; }
            public bool Visible => true;
            public bool Temporary { get; set; }
            public bool CopiedFromSource { get; set; }
            public Button Copy(object target, int before)
            {
                var copied = ((Bar)target).Controls.Add(Type, Id, System.Type.Missing, before, true);
                copied.Caption = Caption; copied.BuiltIn = BuiltIn; copied.CopiedFromSource = true;
                ((Bar)target).Controls.Copied?.Invoke(copied); return copied;
            }
            public void Delete() {if(FailDelete)throw new InvalidOperationException("native delete rejected");if(!IgnoreDelete)Owner.Remove(this);}
        }
    }
}
