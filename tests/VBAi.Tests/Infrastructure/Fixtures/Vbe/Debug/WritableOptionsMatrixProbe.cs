namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;
    internal sealed class WritableOptionsMatrixProbe : VbeDebugWindows.IWritableOptionsProbe
    {
        internal readonly List<string> Names=new List<string>{"Editor"};
        internal readonly List<VbeDebugWindows.OptionsControl> Items=new List<VbeDebugWindows.OptionsControl>{new VbeDebugWindows.OptionsControl{Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On"}};
        internal bool Open=true, KeepOpen, IgnoreWrite;
        internal int Writes, Accepts, Closes;
        internal Action OnTabs,OnWrite,OnAccept;
        public IntPtr Dialog()=>Open?new IntPtr(1):IntPtr.Zero;
        public IList<string> Tabs(IntPtr dialog){OnTabs?.Invoke();return Names;}
        public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog,int index)=>Items;
        public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog)=>new VbeDebugWindows.OptionsChoice[0];
        public void Close(IntPtr dialog){Closes++;Open=false;}
        public void Pause(int milliseconds){}
        public void Write(IntPtr dialog,int index,string name,string type,object value)
        {Writes++;OnWrite?.Invoke();if(!IgnoreWrite)Items.Single(x=>x.Name==name && x.Type==type).Value=type=="ControlType.CheckBox"?((bool)value?"On":"Off"):type=="ControlType.RadioButton"?(object)true:Convert.ToString(value);}
        public void Accept(IntPtr dialog){Accepts++;OnAccept?.Invoke();if(!KeepOpen)Open=false;}
        internal Request Request()
        {dynamic read=VbeDebugWindows.ReadVbeOptions(this);Open=true;Closes=0;return new Request{Pane=Names[0],Property=Items[0].Name,Value=false,ExpectedOptionsVersion=read.OptionsVersion};}
    }
}
