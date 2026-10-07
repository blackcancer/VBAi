using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    public sealed partial class BridgeServerTests
    {
        [STATestMethod, DoNotParallelize]
        public void PathDiagnosticIsDisabledWithoutHostOptInAndRejectsAdditionalFields()
        {
            string original = Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName, null);
                using (var dispatcher = new Control())
                {
                    var handle = dispatcher.Handle;
                    int pid = (int)PathVisibilityObservation.ProcessId;
                    using (var server = new BridgeServer(dispatcher, null, pid))
                    {
                        Environment.SetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName, @"C:\never-read.visibility.json");
                        server.Execute = _ => { Assert.Fail("The internal diagnostic must not invoke VBE session commands."); return null; };
                        server.Start();
                        var disabled = SendWithMessagePump(pid, "{\"Command\":\"diagnostic_path_visibility\"}");
                        Assert.AreEqual(false, disabled["Ok"]); StringAssert.Contains((string)disabled["Error"], "disabled");
                        var injected = SendWithMessagePump(pid, "{\"Command\":\"diagnostic_path_visibility\",\"Path\":\"private\"}");
                        Assert.AreEqual(false, injected["Ok"]); StringAssert.Contains((string)injected["Error"], "Only the fixed");
                    }
                }
            }
            finally { Environment.SetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName, original); }
        }

        [STATestMethod, DoNotParallelize]
        public void FixedAllowlistDiagnosticRunsOnCapturedOwnerStaAndReadsNoFileContents()
        {
            string original = Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName);
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Guid.NewGuid().ToString("N"));
            string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string manifest = Path.Combine(temp, Guid.NewGuid().ToString("N") + ".visibility.json");
            Directory.CreateDirectory(local); Directory.CreateDirectory(temp);
            try
            {
                File.WriteAllText(manifest, new JavaScriptSerializer().Serialize(new { Version = 1, LocalAppData = local, Temp = temp }), new System.Text.UTF8Encoding(false));
                Environment.SetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName, manifest);
                using (var dispatcher = new Control())
                {
                    var handle = dispatcher.Handle;
                    int pid = (int)PathVisibilityObservation.ProcessId;
                    uint nativeTid = PathVisibilityObservation.ThreadId;
                    using (var server = new BridgeServer(dispatcher, null, pid))
                    {
                        // A post-connection manifest change cannot turn the captured allowlist into a free-path gateway.
                        File.WriteAllText(manifest, "{\"Version\":1,\"LocalAppData\":\"C:\\\\private\",\"Temp\":\"C:\\\\private\"}");
                        server.Execute = _ => { Assert.Fail("Readonly diagnostic must not enter native session execution."); return null; };
                        server.Start();
                        var response = SendWithMessagePump(pid, "{\"Command\":\"diagnostic_path_visibility\"}");
                        Assert.AreEqual(true, response["Ok"], new JavaScriptSerializer().Serialize(response));
                        var data = (System.Collections.Generic.IDictionary<string, object>)response["Data"];
                        Assert.AreEqual(pid, Convert.ToInt32(data["ProcessId"])); Assert.AreEqual(nativeTid, Convert.ToUInt32(data["NativeThreadId"]));
                        Assert.AreEqual("STA", data["Apartment"]); Assert.AreEqual(Thread.CurrentThread.ManagedThreadId, data["ManagedThreadId"]);
                        var rows = (object[])data["Paths"];
                        Assert.AreEqual(local, ((System.Collections.Generic.IDictionary<string, object>)rows[0])["Path"]);
                        Assert.AreEqual(temp, ((System.Collections.Generic.IDictionary<string, object>)rows[2])["Path"]);
                        var token = (System.Collections.Generic.IDictionary<string, object>)data["EffectiveTokenBefore"];
                        Assert.AreEqual("READ", token["State"]);
                        Assert.IsTrue(token.ContainsKey("UserSid") && token.ContainsKey("IntegritySid") && token.ContainsKey("AuthenticationId") && token.ContainsKey("TokenId"));
                        Assert.IsFalse(File.Exists(Path.Combine(local, PathVisibilityDiagnostic.SyntheticName)), "Missing synthetic file must be observed without creating it.");
                    }
                }
            }
            finally
            {
                Environment.SetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName, original);
                Directory.Delete(local, true); Directory.Delete(temp, true);
            }
        }
    }
}
