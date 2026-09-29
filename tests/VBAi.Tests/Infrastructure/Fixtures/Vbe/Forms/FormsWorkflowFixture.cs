namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Windows.Forms;
    using VBAi;

    // Live Designer contracts and clipboard serialization; no orchestration method is replaced.
    public sealed class FormsWorkflowFixture : IDisposable
    {
        internal readonly Host Vbe=new Host();
        internal readonly WorkflowProject Project;
        internal readonly WorkflowForm Form;
        internal readonly VbeForms Service;
        private readonly Dictionary<FieldInfo,object> saved=new Dictionary<FieldInfo,object>();
        private readonly List<DesignerNode> nodes=new List<DesignerNode>();
        private readonly Dictionary<string,FormLayoutBox[]> payloads=new Dictionary<string,FormLayoutBox[]>();
        internal uint Sequence=10;
        internal IDataObject ClipboardData=new DataObject();
        internal Action OnReadClipboard,OnWriteClipboard;
        internal FormsWorkflowFixture()
        {
            foreach(string name in new[]{"DesignerClipboardSequence","ReadDesignerClipboard","WriteDesignerClipboard"})
            {var field=typeof(VbeForms).GetField(name,BindingFlags.Static|BindingFlags.NonPublic);saved.Add(field,field.GetValue(null));}
            VbeForms.DesignerClipboardSequence=()=>Sequence;
            VbeForms.ReadDesignerClipboard=()=>{OnReadClipboard?.Invoke();return ClipboardData;};
            VbeForms.WriteDesignerClipboard=(data,persist)=>{OnWriteClipboard?.Invoke();ClipboardData=data;Sequence++;};
            Project=new WorkflowProject(); Form=new WorkflowForm(this);Project.VBComponents.Add(Form);Vbe.VBProjects.Add(Project);
            Vbe.ActiveVBProject=Project;Vbe.ActiveWindow=Form.Window;
            Form.Window.OnFocus=()=>{Vbe.ActiveVBProject=Project;Vbe.ActiveWindow=Form.Window;};
            Form.Designer.Controls.Add("Label","A");Form.Designer.Controls.Add("Label","B").Left=40;
            Service=new VbeForms(Vbe);
        }
        internal Request Request(string action=null,string parent=null)
        {
            dynamic tree=Service.Tree(Project.Name,Form.Name);
            return new Request{Project=Project.Name,Form=Form.Name,Action=action,ParentPath=parent,ExpectedTreeVersion=tree.TreeVersion,
                ExpectedMode=2,ExpectedSha256=EditorDebugFixture.Hash(Form.CodeModule.Code),ControlCaption="Run"};
        }
        internal Request ClipboardRequest(string action=null,string parent=null)
        {
            var request=Request(action,parent);dynamic state=Service.ClipboardState(Project.Name,Form.Name,parent);
            request.ExpectedDesignerSelectionVersion=state.SelectionVersion;request.ExpectedClipboardVersion=state.ClipboardVersion;return request;
        }
        internal void Copy(DesignerNode container)
        {
            var boxes=container.Controls.Items.Where(x=>x.InSelection).Select(x=>new FormLayoutBox{Path=x.Name,Left=x.Left,Top=x.Top,Width=x.Width,Height=x.Height}).ToArray();
            Sequence++;var bytes=BitConverter.GetBytes(Sequence);payloads[BitConverter.ToString(bytes)]=boxes;
            var data=new DataObject();data.SetData("MS Forms Bag",false,new MemoryStream(bytes));ClipboardData=data;
        }
        internal void Cut(DesignerNode container)
        {Sequence++;container.Controls.Items.RemoveAll(x=>x.InSelection);}
        internal void Paste(DesignerNode container)
        {
            var stream=(MemoryStream)ClipboardData.GetData("MS Forms Bag",false);
            foreach(var box in payloads[BitConverter.ToString(stream.ToArray())])
            {
                string name=box.Path;
                if(container.Controls.Items.Any(x=>x.Name==name))name+="Copy"+container.Controls.Count;
                var node=container.Controls.Add("Label",name);node.Left=box.Left+10;node.Top=box.Top+10;node.Width=box.Width;node.Height=box.Height;
            }
        }
        public void Dispose()
        {try{foreach(var pair in saved)pair.Key.SetValue(null,pair.Value);}finally{foreach(var node in nodes)node.DisposeProvider();}}
        public sealed class Host
        {
            public List<WorkflowProject> VBProjects{get;}=new List<WorkflowProject>();
            public object ActiveVBProject{get;set;}
            public object ActiveWindow{get;set;}
            public Bars CommandBars{get;}=new Bars();
        }
        public sealed class WorkflowProject
        {
            public string Name{get;set;}="FormsCoverage";
            public int Mode{get;set;}=2;
            public List<WorkflowForm> VBComponents{get;}=new List<WorkflowForm>();
        }
        public sealed class WorkflowForm
        {
            private readonly DesignerNode designer;
            internal bool FailDesigner;
            public string Name{get;set;}="Form1";
            public int Type=>3;
            public DesignerNode Designer{get{if(FailDesigner)throw new InvalidOperationException("designer unavailable");return designer;}}
            public VbeFormsInitializerTests.FakeCodeModule CodeModule{get;}=new VbeFormsInitializerTests.FakeCodeModule("Sub Example()\r\nEnd Sub");
            public Window Window{get;}=new Window();
            public Window DesignerWindow(){return Window;}
            public Property[] Properties{get;}
            internal WorkflowForm(FormsWorkflowFixture owner)
            {
                designer=new DesignerNode(owner,"UserForm","Form1",null);
                Properties=new[]{new Property("Caption",()=>designer.Caption),new Property("CanUndo",()=>designer.CanUndo),new Property("CanRedo",()=>designer.CanRedo)};
            }
        }
        public sealed class Property
        {
            private readonly Func<object> read;
            public string Name{get;}
            public int NumIndices=>0;
            public object Value=>read();
            internal Property(string name,Func<object> value){Name=name;read=value;}
        }
        public sealed class Window
        {
            public bool Visible{get;set;}
            internal Action OnFocus;
            public void SetFocus(){OnFocus?.Invoke();}
        }
        public sealed class Bars
        {
            public object Control{get;set;}=new RunControl();
            public object FindControl(int type,int id){return Control;}
        }
        public sealed class RunControl
        {
            public int Id{get;set;}=186;public bool Enabled{get;set;}=true;public string Caption{get;set;}="Run";
            internal Action OnExecute;internal int Calls;
            public void Execute(){Calls++;OnExecute?.Invoke();}
        }
        public sealed class NodeCollection:IEnumerable<DesignerNode>
        {
            internal readonly List<DesignerNode> Items=new List<DesignerNode>();
            private readonly DesignerNode parent;
            internal NodeCollection(DesignerNode owner){parent=owner;}
            public int Count=>Items.Count;
            public DesignerNode Add(string type,string name)
            {var node=new DesignerNode(parent.Scene,type,name,parent);node.tab=Items.Count;Items.Add(node);return node;}
            public DesignerNode Item(string name){return Items.Single(x=>x.Name==name);}
            public IEnumerator<DesignerNode> GetEnumerator(){return Items.ToList().GetEnumerator();}
            IEnumerator IEnumerable.GetEnumerator(){return GetEnumerator();}
        }
        public sealed class DesignerNode
        {
            internal readonly FormsWorkflowFixture Scene;
            private readonly TypeDescriptionProvider provider;
            internal Action OnCopy,OnCut,OnPaste,OnUndo,OnRedo;
            internal Action<string,double> OnSetGeometry;
            internal Func<string,double,double> OnReadGeometry;
            internal Action<int> OnSetTab;
            internal Func<int,int> OnReadTab;
            internal Action<bool> OnSelection;
            internal bool IgnoreSelection,IgnoreTab;
            internal int tab;
            private bool selected;
            private double left,top,width=20,height=10;
            public string Name{get;set;}
            public object Parent{get;}
            public NodeCollection Controls{get;}
            public NodeCollection Pages{get;}
            public NodeCollection Tabs{get;}
            public string Caption{get;set;}="Caption";
            public double InsideWidth{get;set;}=300;
            public double InsideHeight{get;set;}=200;
            public bool CanUndo{get;set;}=true;
            public bool CanRedo{get;set;}=true;
            public bool CanPaste{get;set;}=true;
            public IEnumerable<DesignerNode> Selected=>Controls.Items.Where(x=>x.InSelection).ToArray();
            public bool InSelection{get{return selected;}set{OnSelection?.Invoke(value);if(!IgnoreSelection)selected=value;}}
            private double Read(string name,double value){return OnReadGeometry==null?value:OnReadGeometry(name,value);}
            public double Left{get{return Read("Left",left);}set{OnSetGeometry?.Invoke("Left",value);left=value;}}
            public double Top{get{return Read("Top",top);}set{OnSetGeometry?.Invoke("Top",value);top=value;}}
            public double Width{get{return Read("Width",width);}set{OnSetGeometry?.Invoke("Width",value);width=value;}}
            public double Height{get{return Read("Height",height);}set{OnSetGeometry?.Invoke("Height",value);height=value;}}
            public int TabIndex
            {
                get{return OnReadTab==null?tab:OnReadTab(tab);}
                set{OnSetTab?.Invoke(value);if(IgnoreTab)return;var siblings=Parent is DesignerNode parent?parent.Controls.Items:new List<DesignerNode>();
                    foreach(var sibling in siblings.Where(x=>x!=this)) {if(value<tab&&sibling.tab>=value&&sibling.tab<tab)sibling.tab++;if(value>tab&&sibling.tab>tab&&sibling.tab<=value)sibling.tab--;}
                    tab=value;}
            }
            internal DesignerNode(FormsWorkflowFixture scene,string kind,string name,object parent)
            {
                Scene=scene;Name=name;Parent=parent;Controls=new NodeCollection(this);Pages=new NodeCollection(this);Tabs=new NodeCollection(this);
                provider=new NamedProvider(TypeDescriptor.GetProvider(this),kind);TypeDescriptor.AddProvider(provider,this);scene.nodes.Add(this);
            }
            internal void DisposeProvider(){TypeDescriptor.RemoveProvider(provider,this);}
            public void Copy(){if(OnCopy!=null)OnCopy();else Scene.Copy(this);}
            public void Cut(){if(OnCut!=null)OnCut();else Scene.Cut(this);}
            public void Paste(){if(OnPaste!=null)OnPaste();else Scene.Paste(this);}
            public void UndoAction(){if(OnUndo!=null)OnUndo();else Caption+=" undo";}
            public void RedoAction(){if(OnRedo!=null)OnRedo();else Caption+=" redo";}
        }
        private sealed class NamedProvider:TypeDescriptionProvider
        {
            private readonly TypeDescriptionProvider parent;private readonly string name;
            internal NamedProvider(TypeDescriptionProvider source,string kind){parent=source;name=kind;}
            public override ICustomTypeDescriptor GetTypeDescriptor(Type type,object instance){return new NamedDescriptor(parent.GetTypeDescriptor(type,instance),name);}
        }
        private sealed class NamedDescriptor:CustomTypeDescriptor
        {
            private readonly string name;internal NamedDescriptor(ICustomTypeDescriptor source,string kind):base(source){name=kind;}
            public override string GetClassName(){return name;}
        }
    }
}
