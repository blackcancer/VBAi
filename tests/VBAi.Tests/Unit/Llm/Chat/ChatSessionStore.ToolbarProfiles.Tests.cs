namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatSessionStoreTests
    {
        [TestMethod]
        public void ToolbarStoreNormalizesKeysReplacesAndRollsBackAllValidationFailures()
        {
            using (var scope = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(scope.Root,"profiles.sqlite")))
            {
                Action<VbeToolbarProfiles.Bar[]> validate = bars => { foreach (var bar in bars) if (bar == null || bar.Name == "Bad") throw new InvalidOperationException("validation rejected"); };
                Assert.AreEqual(0,store.ReadToolbarProfiles().Length);
                store.UpdateToolbarProfile("VBAi - One",StoredBar(),validate);
                store.UpdateToolbarProfile("vbai - one",new VbeToolbarProfiles.Bar { Name="VBAi - One",Position=4,Commands=new VbeToolbarProfiles.Command[0] },validate);
                store.UpdateToolbarProfile("VBAi - Two",StoredBar("VBAi - Two"),validate);
                Assert.AreEqual(2,store.ReadToolbarProfiles().Length); Assert.AreEqual(4,store.ReadToolbarProfiles()[0].Position);
                Assert.ThrowsException<InvalidOperationException>(()=>store.UpdateToolbarProfile("VBAi - One",StoredBar("VBAi - Different"),validate));
                Assert.ThrowsException<InvalidOperationException>(()=>store.UpdateToolbarProfile("Bad",StoredBar("Bad"),validate));
                Assert.ThrowsException<InvalidOperationException>(()=>store.UpdateToolbarProfile("VBAi - One",StoredBar(),bars=>{throw new InvalidOperationException("original rejected");}));
                Assert.AreEqual(2,store.ReadToolbarProfiles().Length);
                var huge=StoredBar(); huge.Commands=new[] { new VbeToolbarProfiles.Command { Id=42,Caption=new string('x',128*1024),Tag="owned" } };
                Assert.ThrowsException<InvalidOperationException>(()=>store.UpdateToolbarProfile(huge.Name,huge,bars=>{}));
                Assert.AreEqual(2,store.ReadToolbarProfiles().Length);
                store.UpdateToolbarProfile("vbai - one",null,validate); Assert.AreEqual("VBAi - Two",store.ReadToolbarProfiles()[0].Name);
                store.UpdateToolbarProfile("VBAi - Missing",null,validate); Assert.AreEqual(1,store.ReadToolbarProfiles().Length);
            }
        }

        [TestMethod]
        public void ToolbarStoreBoundsDeserializationAndLeavesInvalidDatabaseUntouched()
        {
            using(var scope=new LlmBoundaryScope())
            {
                string path=Path.Combine(scope.Root,"profiles.sqlite");
                using(var store=new ChatSessionStore(path))
                {
                    store.ReadToolbarProfiles();
                    Sql(store,"INSERT INTO vbe_toolbar_profiles VALUES(?1,?2)","NULL","null");
                    Assert.IsNull(store.ReadToolbarProfiles()[0]);
                    Assert.ThrowsException<InvalidOperationException>(()=>new VbeToolbarProfiles(path).Read());
                    Sql(store,"DELETE FROM vbe_toolbar_profiles");
                    Sql(store,"INSERT INTO vbe_toolbar_profiles VALUES(?1,?2)","BAD","{");
                    Assert.ThrowsException<ArgumentException>(()=>store.ReadToolbarProfiles());
                    Sql(store,"DELETE FROM vbe_toolbar_profiles");
                    Sql(store,"INSERT INTO vbe_toolbar_profiles VALUES(?1,?2)","BIG",new string(' ',128*1024+1));
                    Assert.ThrowsException<InvalidOperationException>(()=>store.ReadToolbarProfiles());
                    Sql(store,"DELETE FROM vbe_toolbar_profiles");
                    for(int i=0;i<34;i++) Sql(store,"INSERT INTO vbe_toolbar_profiles VALUES(?1,?2)",i.ToString("D2"),"{\"Name\":\"VBAi - "+i+"\",\"Commands\":[]}");
                    Assert.AreEqual(33,store.ReadToolbarProfiles().Length);
                    Assert.ThrowsException<InvalidOperationException>(()=>new VbeToolbarProfiles(path).Read());
                }
            }
        }

        [TestMethod]
        public void ToolbarTransactionsRollbackNativeWriteAndCommitFailuresAndReleaseLocks()
        {
            foreach(string prefix in new[] { "BEGIN IMMEDIATE", "INSERT OR REPLACE", "COMMIT" })
            using(var scope=new LlmBoundaryScope())
            using(var store=new ChatSessionStore(Path.Combine(scope.Root,"profiles.sqlite")))
            {
                store.UpdateToolbarProfile("VBAi - One",StoredBar(),bars=>{});
                var native=store.StepNative; bool failed=false;
                store.StepNative=statement=> {
                    string sql=ToolbarStatementSql(statement);
                    if(!failed && sql.StartsWith(prefix,StringComparison.Ordinal)) { failed=true; return 5; }
                    return native(statement);
                };
                Assert.ThrowsException<IOException>(()=>store.UpdateToolbarProfile("VBAi - Two",StoredBar("VBAi - Two"),bars=>{}));
                Assert.IsTrue(failed,prefix); store.StepNative=native;
                Assert.AreEqual(1,store.ReadToolbarProfiles().Length,prefix);
                store.UpdateToolbarProfile("VBAi - Two",StoredBar("VBAi - Two"),bars=>{});
                Assert.AreEqual(2,store.ReadToolbarProfiles().Length,"Rollback must release the transaction.");
            }
        }
    }
}
