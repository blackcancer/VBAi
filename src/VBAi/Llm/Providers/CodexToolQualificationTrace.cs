using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Opt-in bounded observations of synthetic provider calls; never dispatches a tool.</summary>
    internal sealed class CodexToolQualificationTrace
    {
        internal const string ManifestVariable = "VBAi_CODEX_TOOL_QUALIFICATION_MANIFEST";
        internal const int MaximumRecords = 64;
        internal sealed class Manifest
        {
            public int ProcessId { get; set; }
            public long ProcessBirthUtcTicks { get; set; }
            public string AssemblyMvid { get; set; }
            public string Nonce { get; set; }
            public long ExpiresUtcTicks { get; set; }
            public string ProjectHash { get; set; }
            public string ExpectedModuleHash { get; set; }
            public string MissingModuleHash { get; set; }
        }
        internal sealed class Ticket
        {
            internal string RequestHash, CallHash, ThreadHash, TurnHash, Tool, Target;
            internal int Invocation;
            internal bool Admitted;
        }
        private readonly object gate = new object();
        private readonly Manifest manifest;
        private readonly Func<DateTime> utcNow;
        private readonly Action<string> sink;
        private Action closeSink;
        private Task writer = Task.CompletedTask;
        private int sequence, invocation, bytes;
        private bool failed, closed;
        internal CodexToolQualificationTrace(Manifest manifest, Func<DateTime> utcNow, Action<string> sink)
        { this.manifest = manifest; this.utcNow = utcNow; this.sink = sink; }

        /// <summary>Reads only an explicit small manifest matching this process, assembly and cached scope.</summary>
        internal static CodexToolQualificationTrace TryCreate(string cachedProject)
        {
            try { return TryCreate(cachedProject, Environment.GetEnvironmentVariable(ManifestVariable)); } catch { return null; }
        }
        internal static CodexToolQualificationTrace TryCreate(string cachedProject, string path)
        {
            try
            {
                if (!IsLocalManifestPath(path)) return null;
                path = Path.GetFullPath(path);
                byte[] raw;
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (stream.Length < 2 || stream.Length > 8192) return null;
                    raw = new byte[(int)stream.Length];
                    int read = 0, count;
                    while (read < raw.Length && (count = stream.Read(raw, read, raw.Length-read)) > 0) read += count;
                    if (read != raw.Length || stream.ReadByte() != -1) return null;
                }
                var manifest = new JavaScriptSerializer().Deserialize<Manifest>(new UTF8Encoding(false,true).GetString(raw).TrimStart('\uFEFF'));
                using (var process = Process.GetCurrentProcess())
                    if (!Matches(manifest, process.Id, process.StartTime.ToUniversalTime().Ticks,
                        typeof(CodexToolQualificationTrace).Assembly.ManifestModule.ModuleVersionId, cachedProject, DateTime.UtcNow)) return null;
                string output = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(path)), "tool-receipts-"+manifest.Nonce+".jsonl");
                StreamWriter file = null;
                var trace = new CodexToolQualificationTrace(manifest, () => DateTime.UtcNow, line =>
                {
                    if (file == null) file = new StreamWriter(new FileStream(output,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(false)) { AutoFlush=true };
                    file.WriteLine(line);
                });
                trace.closeSink = () => file?.Dispose();
                trace.Record(null,"Armed",null,Hash(raw));
                return trace;
            }
            catch { return null; }
        }
        /// <summary>Rejects network, drive-relative and reparse paths before synchronous manifest reads.</summary>
        internal static bool IsLocalManifestPath(string path)
        {
            try
            {
                if(string.IsNullOrWhiteSpace(path)||path.Length<3||path[1]!=':'||(path[2]!='\\'&&path[2]!='/'))return false;
                string full=Path.GetFullPath(path), root=Path.GetPathRoot(full);
                if(full.IndexOf(':',2)>=0 || !string.Equals(full,path.Replace('/','\\'),StringComparison.OrdinalIgnoreCase))return false;
                if(new DriveInfo(root).DriveType!=DriveType.Fixed)return false;
                string current=full;int depth=0;
                while(current!=null)
                {
                    if(++depth>32||(File.GetAttributes(current)&FileAttributes.ReparsePoint)!=0)return false;
                    current=Path.GetDirectoryName(current);
                }
                return true;
            }
            catch{return false;}
        }
        internal static bool Matches(Manifest value,int pid,long birth,Guid mvid,string project,DateTime now)
        {
            Guid nonce, declaredMvid;
            return value != null && value.ProcessId==pid && value.ProcessBirthUtcTicks==birth &&
                Guid.TryParse(value.AssemblyMvid,out declaredMvid) && declaredMvid==mvid &&
                Guid.TryParseExact(value.Nonce,"N",out nonce) && nonce!=Guid.Empty &&
                value.ExpiresUtcTicks>now.Ticks && value.ExpiresUtcTicks<=now.AddMinutes(20).Ticks &&
                IsHash(value.ProjectHash) && IsHash(value.ExpectedModuleHash) && IsHash(value.MissingModuleHash) &&
                !string.IsNullOrWhiteSpace(project) && Hash(project)==value.ProjectHash && value.ExpectedModuleHash!=value.MissingModuleHash;
        }
        internal Ticket Received(object request,IDictionary<string,object> parameters,string cachedProject)
        {
            try
            {
                lock(gate) if(failed||closed||sequence>=MaximumRecords||utcNow().Ticks>=manifest.ExpiresUtcTicks)return null;
                if (Hash(cachedProject??"")!=manifest.ProjectHash) return null;
                object raw;
                var arguments=parameters!=null && parameters.TryGetValue("arguments",out raw) ? raw as IDictionary<string,object> : null;
                string project=Text(arguments,"Project"), module=Text(arguments,"Module");
                string target=project==null || module==null || Hash(project)!=manifest.ProjectHash ? "Other" :
                    Hash(module??"")==manifest.ExpectedModuleHash ? "Expected" : Hash(module??"")==manifest.MissingModuleHash ? "Missing" : "Other";
                var ticket=new Ticket { RequestHash=Hash(Convert.ToString(request)), CallHash=Hash(Text(parameters,"callId")??Text(parameters,"itemId")??"tool-"+Convert.ToString(request)),
                    ThreadHash=Hash(Text(parameters,"threadId")??""), TurnHash=Hash(Text(parameters,"turnId")??""),
                    Tool=Text(parameters,"tool")=="read_module" ? "read_module" : "Other", Target=target };
                lock(gate) ticket.Invocation=++invocation;
                Record(ticket,"Received",null);
                return ticket;
            }
            catch { return null; }
        }
        internal void Admitted(Ticket ticket,string cachedProject)
        {
            try
            {
                if(ticket==null)return;
                if(Hash(cachedProject??"")!=manifest.ProjectHash){Record(ticket,"ScopeChanged",null);return;}
                ticket.Admitted=true;Record(ticket,"Admitted",null);
            }
            catch { }
        }
        internal void Returned(Ticket ticket,bool ok) { if(ticket!=null)Record(ticket,"Returned",ok); }
        internal void Rejected(Ticket ticket) { if(ticket!=null)Record(ticket,ticket.Admitted ? "ExceptionAfterAdmission" : "RejectedBeforeAdmission",null); }
        private void Record(Ticket ticket,string stage,bool? ok,string manifestHash=null)
        {
            try
            {
                lock(gate)
                {
                    if(failed||closed||sequence>=MaximumRecords||utcNow().Ticks>=manifest.ExpiresUtcTicks)return;
                    string line=new JavaScriptSerializer().Serialize(new { Sequence=++sequence,Nonce=manifest.Nonce,Stage=stage,
                        Invocation=ticket?.Invocation,RequestHash=ticket?.RequestHash,CallHash=ticket?.CallHash,ThreadHash=ticket?.ThreadHash,
                        TurnHash=ticket?.TurnHash,Tool=ticket?.Tool,Target=ticket?.Target,Ok=ok,ManifestHash=manifestHash });
                    bytes+=Encoding.UTF8.GetByteCount(line)+1;if(bytes>65536){failed=true;return;}
                    writer=writer.ContinueWith(_=> { try { lock(gate){if(failed)return;} sink(line); } catch { lock(gate)failed=true; } },TaskScheduler.Default);
                }
            }
            catch { }
        }
        /// <summary>Queues a terminal marker and disposes the writer without waiting on the owner thread.</summary>
        internal void Close()
        {
            try
            {
                Record(null,"Closed",null);
                lock(gate)
                {
                    if(closed)return;closed=true;
                    writer=writer.ContinueWith(_=>{try{closeSink?.Invoke();}catch{}},TaskScheduler.Default);
                }
            }
            catch { }
        }
        internal Task PendingWrites { get { lock(gate)return writer; } }
        internal static string Hash(string text) => Hash(Encoding.UTF8.GetBytes(text??""));
        private static string Hash(byte[] raw)
        { using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(raw)).Replace("-",""); }
        private static bool IsHash(string value) => value!=null && value.Length==64 && value.All(c=>c>='0'&&c<='9'||c>='A'&&c<='F');
        private static string Text(IDictionary<string,object> values,string key)
        {object raw;return values!=null && values.TryGetValue(key,out raw) && raw is string ? (string)raw : null;}
    }
}