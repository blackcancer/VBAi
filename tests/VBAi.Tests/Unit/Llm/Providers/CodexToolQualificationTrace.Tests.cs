namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class CodexToolQualificationTraceTests
    {
        internal static CodexToolQualificationTrace.Manifest Manifest(DateTime now)
        {
            return new CodexToolQualificationTrace.Manifest { ProcessId=17,ProcessBirthUtcTicks=23,
                AssemblyMvid=Guid.Empty.ToString(),Nonce=Guid.NewGuid().ToString("N"),ExpiresUtcTicks=now.AddMinutes(5).Ticks,
                ProjectHash=CodexToolQualificationTrace.Hash("SyntheticProject"),ExpectedModuleHash=CodexToolQualificationTrace.Hash("SyntheticExpected"),
                MissingModuleHash=CodexToolQualificationTrace.Hash("SyntheticMissing") };
        }
        internal static Dictionary<string,object> Parameters(string module="SyntheticExpected",string project="SyntheticProject") => new Dictionary<string,object> {
            {"threadId","private-thread"},{"turnId","private-turn"},{"callId","private-call"},{"tool","read_module"},
            {"arguments",new Dictionary<string,object>{{"Project",project},{"Module",module}}} };

        [TestMethod]
        public void ReadModuleObservationArgumentsMatchTheActualRequiredSchema()
        {
            var json=new JavaScriptSerializer();
            var definition=LlmVbeTools.Definitions.Select(d=>(Dictionary<string,object>)json.DeserializeObject(json.Serialize(d)))
                .Select(d=>(Dictionary<string,object>)d["function"]).Single(d=>Equals(d["name"],"read_module"));
            var schema=(Dictionary<string,object>)definition["parameters"];
            var properties=(Dictionary<string,object>)schema["properties"];
            CollectionAssert.AreEquivalent(new[]{"Project","Module"},properties.Keys.ToArray());
            CollectionAssert.AreEquivalent(new[]{"Project","Module"},((object[])schema["required"]).Cast<string>().ToArray());
            Assert.AreEqual(false,schema["additionalProperties"]);
            CollectionAssert.AreEquivalent(properties.Keys.ToArray(),((Dictionary<string,object>)Parameters()["arguments"]).Keys.ToArray());
        }

        [DataTestMethod]
        [DataRow("SyntheticProject","SyntheticExpected","Expected")]
        [DataRow("SyntheticProject","SyntheticMissing","Missing")]
        [DataRow("OtherProject","SyntheticExpected","Other")]
        [DataRow("OtherProject","SyntheticMissing","Other")]
        [DataRow("SyntheticProject","OtherModule","Other")]
        [DataRow("syntheticProject","SyntheticExpected","Other")]
        [DataRow("SyntheticProject","syntheticExpected","Other")]
        [DataRow(null,"SyntheticExpected","Other")]
        [DataRow("SyntheticProject",null,"Other")]
        public async Task CanonicalSchemaTargetsRemainExactAfterJsonTransport(string project,string module,string expectedTarget)
        {
            var now=DateTime.UtcNow;var lines=new List<string>();var json=new JavaScriptSerializer();
            var trace=new CodexToolQualificationTrace(Manifest(now),()=>now,lines.Add);
            var parameters=(Dictionary<string,object>)json.DeserializeObject(json.Serialize(Parameters(module,project)));
            var ticket=trace.Received("synthetic-request",parameters,"SyntheticProject");
            Assert.IsNotNull(ticket);Assert.AreEqual(expectedTarget,ticket.Target);
            trace.Admitted(ticket,"SyntheticProject");trace.Returned(ticket,true);trace.Close();await trace.PendingWrites;
            var records=lines.Select(line=>(Dictionary<string,object>)json.DeserializeObject(line)).ToArray();
            foreach(var record in records.Where(r=>!Equals(r["Stage"],"Closed")))Assert.AreEqual(expectedTarget,record["Target"]);
            Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Received")));
            Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Admitted")));
            Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Returned")&&Equals(r["Ok"],true)));
        }

        [DataTestMethod]
        [DataRow("project","module")]
        [DataRow("project","Module")]
        [DataRow("Project","module")]
        public async Task NoncanonicalArgumentNamesCannotQualifyAnExpectedTarget(string projectKey,string moduleKey)
        {
            var now=DateTime.UtcNow;var trace=new CodexToolQualificationTrace(Manifest(now),()=>now,_=>{});
            var parameters=Parameters();parameters["arguments"]=new Dictionary<string,object>{{projectKey,"SyntheticProject"},{moduleKey,"SyntheticExpected"}};
            var ticket=trace.Received("synthetic-request",parameters,"SyntheticProject");
            Assert.IsNotNull(ticket);Assert.AreEqual("Other",ticket.Target);
            trace.Close();await trace.PendingWrites;
        }

        [DataTestMethod]
        [DataRow("Project")]
        [DataRow("Module")]
        public async Task NonstringRequiredArgumentsCannotQualifyAnExpectedTarget(string field)
        {
            var now=DateTime.UtcNow;var trace=new CodexToolQualificationTrace(Manifest(now),()=>now,_=>{});
            var parameters=Parameters();((Dictionary<string,object>)parameters["arguments"])[field]=17;
            var ticket=trace.Received("synthetic-request",parameters,"SyntheticProject");
            Assert.IsNotNull(ticket);Assert.AreEqual("Other",ticket.Target);
            trace.Close();await trace.PendingWrites;
        }

        [TestMethod]
        public void ManifestRejectsWrongGenerationAssemblyScopeExpiryAndMalformedHashes()
        {
            var now=DateTime.UtcNow;var manifest=Manifest(now);
            Assert.IsTrue(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
            Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,18,23,Guid.Empty,"SyntheticProject",now));
            Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,24,Guid.Empty,"SyntheticProject",now));
            Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.NewGuid(),"SyntheticProject",now));
            Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"Other",now));
            Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,null,now));
            manifest.ExpiresUtcTicks=now.Ticks;Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
            manifest.ExpiresUtcTicks=now.AddMinutes(21).Ticks;Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
            manifest=Manifest(now);manifest.Nonce="../output";Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
            manifest=Manifest(now);manifest.ExpectedModuleHash=manifest.MissingModuleHash;Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
            manifest=Manifest(now);manifest.ProjectHash="short";Assert.IsFalse(CodexToolQualificationTrace.Matches(manifest,17,23,Guid.Empty,"SyntheticProject",now));
        }
        [TestMethod]
        public async Task DuplicateRequestsRemainSeparateAndRawValuesNeverEnterReceipts()
        {
            var now=DateTime.UtcNow;var lines=new List<string>();var trace=new CodexToolQualificationTrace(Manifest(now),()=>now,lines.Add);
            foreach(string module in new[]{"SyntheticExpected","SyntheticMissing"})
            {
                var parameters=Parameters(module);((Dictionary<string,object>)parameters["arguments"])["secret"]="do-not-persist";
                var ticket=trace.Received("same-rpc-id",parameters,"SyntheticProject");trace.Admitted(ticket,"SyntheticProject");trace.Returned(ticket,module=="SyntheticExpected");
            }
            trace.Close();await trace.PendingWrites;
            var json=new JavaScriptSerializer();var records=lines.Select(l=>json.DeserializeObject(l) as Dictionary<string,object>).ToArray();
            Assert.AreEqual(7,records.Length);Assert.AreEqual(2,records.Count(r=>Equals(r["Stage"],"Received")));
            Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Returned") && Equals(r["Target"],"Expected") && Equals(r["Ok"],true)));
            Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Returned") && Equals(r["Target"],"Missing") && Equals(r["Ok"],false)));
            Assert.AreEqual(records[0]["RequestHash"],records[3]["RequestHash"]);Assert.AreNotEqual(records[0]["Invocation"],records[3]["Invocation"]);
            string output=string.Join("\n",lines);foreach(string secret in new[]{"private-thread","private-turn","private-call","same-rpc-id","SyntheticProject","SyntheticExpected","SyntheticMissing","do-not-persist"})Assert.IsFalse(output.Contains(secret),secret);
        }
        [TestMethod]
        public async Task BoundsExpiryScopeAndWriterFailureCannotThrowIntoToolDispatch()
        {
            var now=DateTime.UtcNow;int writes=0;var trace=new CodexToolQualificationTrace(Manifest(now),()=>now,_=>writes++);
            for(int i=0;i<100;i++){var ticket=trace.Received(i,Parameters(),"SyntheticProject");trace.Admitted(ticket,"SyntheticProject");trace.Returned(ticket,true);}
            trace.Close();await trace.PendingWrites;Assert.AreEqual(CodexToolQualificationTrace.MaximumRecords,writes);
            trace=new CodexToolQualificationTrace(Manifest(now),()=>now.AddHours(1),_=>Assert.Fail("Expired trace wrote"));Assert.IsNull(trace.Received(1,Parameters(),"SyntheticProject"));trace.Close();await trace.PendingWrites;
            trace=new CodexToolQualificationTrace(Manifest(now),()=>now,_=>throw new IOException("synthetic sink failure"));
            var call=trace.Received(1,Parameters(),"SyntheticProject");trace.Admitted(call,"SyntheticProject");trace.Returned(call,true);await trace.PendingWrites;
            Assert.IsNull(trace.Received(2,Parameters(),"SyntheticProject"));trace.Close();await trace.PendingWrites;
            trace=new CodexToolQualificationTrace(Manifest(now),()=>now,_=>Assert.Fail("Foreign scope wrote"));Assert.IsNull(trace.Received(1,Parameters(),"Other"));await trace.PendingWrites;
        }
        [TestMethod]
        public async Task LocalManifestCreatesArmedReceiptAndNeverOverwritesExistingEvidence()
        {
            string directory=Path.Combine(Path.GetTempPath(),"VBAi-tool-receipts-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
            try
            {
                string path=Path.Combine(directory,"manifest.json");var now=DateTime.UtcNow;var manifest=Manifest(now);
                using(var process=System.Diagnostics.Process.GetCurrentProcess()){manifest.ProcessId=process.Id;manifest.ProcessBirthUtcTicks=process.StartTime.ToUniversalTime().Ticks;}
                manifest.AssemblyMvid=typeof(CodexToolQualificationTrace).Assembly.ManifestModule.ModuleVersionId.ToString();
                File.WriteAllText(path,new JavaScriptSerializer().Serialize(manifest),new System.Text.UTF8Encoding(true));
                Assert.IsTrue(CodexToolQualificationTrace.IsLocalManifestPath(path));
                Assert.IsNull(CodexToolQualificationTrace.TryCreate("SyntheticProject",null));
                Assert.IsFalse(CodexToolQualificationTrace.IsLocalManifestPath(path+":stream"));
                Assert.IsFalse(CodexToolQualificationTrace.IsLocalManifestPath(Path.Combine(directory,"..",Path.GetFileName(directory),"manifest.json")));
                Assert.IsFalse(CodexToolQualificationTrace.IsLocalManifestPath("C:relative.json"));
                Assert.IsFalse(CodexToolQualificationTrace.IsLocalManifestPath("//server/share/manifest.json"));
                Assert.IsNull(CodexToolQualificationTrace.TryCreate("Other",path));
                var trace=CodexToolQualificationTrace.TryCreate("SyntheticProject",path);Assert.IsNotNull(trace);trace.Close();await trace.PendingWrites;
                string output=Path.Combine(directory,"tool-receipts-"+manifest.Nonce+".jsonl"), original=File.ReadAllText(output);
                StringAssert.Contains(original,"Armed");StringAssert.Contains(original,"ManifestHash");StringAssert.Contains(original,"Closed");
                trace=CodexToolQualificationTrace.TryCreate("SyntheticProject",path);Assert.IsNotNull(trace);trace.Close();await trace.PendingWrites;
                Assert.AreEqual(original,File.ReadAllText(output));
                File.WriteAllText(path,new string('x',8193));Assert.IsNull(CodexToolQualificationTrace.TryCreate("SyntheticProject",path));
            }
            finally{Directory.Delete(directory,true);}
        }
    }
}