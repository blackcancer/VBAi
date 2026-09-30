using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void ProjectAccessMigrationHandlesEveryLegacyContextAndNoContext()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyCodexWindow(new ChatSessionState())) {
                Set(window,"currentSession",null);Call(window,"MigrateProviderPrivacy");
                foreach(int scenario in Enumerable.Range(0,5)) {
                    var state=new ChatSessionState{ReadAccessPolicyVersion=0};
                    Get<List<ChatEntry>>(window,"transcriptEntries").Clear();Get<List<object>>(window,"messages").Clear();
                    if(scenario==0) Get<List<ChatEntry>>(window,"transcriptEntries").Add(new ChatEntry{Text="local"});
                    if(scenario==1) Get<List<object>>(window,"messages").Add(new{role="user",content="old"});
                    if(scenario==2) { state.CodexThreadId="old"; state.CodexDeveloperInstructionsHash="old-hash"; }
                    if(scenario==3) state.ResumeContext="old";
                    Set(window,"currentSession",state);
                    Call(window,"MigrateProviderPrivacy");
                    Assert.AreEqual(1,state.ReadAccessPolicyVersion);
                    Assert.IsNull(state.CodexThreadId);Assert.IsNull(state.CodexDeveloperInstructionsHash);Assert.IsNull(state.ResumeContext);
                    Assert.AreEqual(scenario==4?0:scenario==0?2:1,state.ProviderHistoryStartIndex);
                    Call(window,"MigrateProviderPrivacy");
                }
                Set(window,"currentSession",null);
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void ProjectAccessDialogFiltersBoundAndAmbiguousProjectsAndAppliesOnlyAcceptedChoices()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyCodexWindow(new ChatSessionState{Scope=@"C:\Temp\P.xlsm"})) {
                AddScope(window,@"C:\Temp\P.xlsm");Get<ComboBox>(window,"scopePicker").SelectedIndex=0;
                Set(window,"scopeSession",runtime.Session);
                var old=Get<ChatSessionState>(window,"currentSession");old.ReadProjectGrants=new[]{"B"};old.SharedContextReadAllowed=true;
                runtime.Host=r=>Response.Success(new object[]{new{Name="P",FileName=@"C:\Temp\P.xlsm"},new{Name="B",FileName=""},new{Name="Dup",FileName=""},new{Name="Dup",FileName=""},new{Name="Saved",FileName=@"C:\S.xlsm"},4});
                int shown=0;
                ChatWindow.ShowModal=(dialog,owner)=>{
                    shown++;Assert.AreSame(window,owner);
                    var access=(ProjectAccessWindow)dialog;
                    var list=(CheckedListBox)access.Controls["projectList"];
                    Assert.AreEqual(2,list.Items.Count);
                    CollectionAssert.AreEqual(new[]{"B"},access.SelectedProjects);Assert.IsTrue(access.SharedContext);
                    return DialogResult.Cancel;
                };
                Call(window,"ConfigureProjectAccess");Assert.AreEqual(1,shown);Assert.AreSame(old,Get<ChatSessionState>(window,"currentSession"));
                ChatWindow.ShowModal=(dialog,owner)=>{
                    var list=(CheckedListBox)dialog.Controls["projectList"];list.SetItemChecked(1,true);
                    ((CheckBox)dialog.Controls["sharedContext"]).Checked=false;return DialogResult.OK;
                };
                Call(window,"ConfigureProjectAccess");
                var current=Get<ChatSessionState>(window,"currentSession");Assert.AreNotSame(old,current);
                CollectionAssert.AreEqual(new[]{"B",@"C:\S.xlsm"},current.ReadProjectGrants);Assert.IsFalse(current.SharedContextReadAllowed);
                Set(window,"currentSession",null);
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void ProjectAccessReportsScopeInventoryAndEmptyDataErrorsWithoutShowingUnownedDialog()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyCodexWindow(new ChatSessionState())) {
                int shown=0;ChatWindow.ShowModal=(d,o)=>{shown++;return DialogResult.Cancel;};
                Set(window,"busy",true);Call(window,"ConfigureProjectAccess");Set(window,"busy",false);
                var state=Get<ChatSessionState>(window,"currentSession");Set(window,"currentSession",null);Call(window,"ConfigureProjectAccess");
                Set(window,"currentSession",state);var tools=Get<LlmVbeTools>(window,"tools");Set(window,"tools",null);Call(window,"ConfigureProjectAccess");Set(window,"tools",tools);
                runtime.Host=r=>Response.Failure("inventory failed");Call(window,"ConfigureProjectAccess");Assert.AreEqual(0,shown);
                runtime.Host=r=>Response.Success(new{});Call(window,"ConfigureProjectAccess");Assert.AreEqual(1,shown);
                Set(window,"scopeSession",runtime.Session);Call(window,"ConfigureProjectAccess");Assert.AreEqual(1,shown);
                Set(window,"currentSession",null);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void PrivacyMigrationDisposesTheOwnedProviderClient()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyCodexWindow(new ChatSessionState{ReadAccessPolicyVersion=0})) {
                Set(window,"codex",Call(window,"CreateCodexClient"));
                Call(window,"MigrateProviderPrivacy");
                Assert.IsNull(Get<CodexAppServerClient>(window,"codex"));
                Set(window,"currentSession",null);
            }
        }
    }
}
