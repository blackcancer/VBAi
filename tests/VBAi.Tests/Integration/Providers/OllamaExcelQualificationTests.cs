using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Live local model to production read-only tools to a disposable native Excel project.</summary>
    [TestClass, TestCategory("OllamaExcel"), DoNotParallelize]
    public sealed class OllamaExcelQualificationTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void LocalModelReadsUnpromptedMarkerFromRealExcelWithoutChangingCode()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OLLAMA_EXCEL_TESTS=1 for live loopback Ollama and disposable native Excel.");
            ExcelScenarioLifetime.Run(lifetime => RunScenario(lifetime, TestContext));
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void RunScenario(ExcelScenarioLifetime lifetime, TestContext context)
        {
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            var existing = Process.GetProcessesByName("EXCEL");
            try { if (existing.Length != 0) Assert.Inconclusive("An existing Excel session must be preserved."); }
            finally { foreach (var process in existing) process.Dispose(); }
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is not installed.");

            const string model = "qwen2.5:3b";
            const string endpoint = "http://127.0.0.1:11434/v1/chat/completions";
            const string moduleName = "OllamaReadFixture";
            string projectName = "OllamaFixture" + Guid.NewGuid().ToString("N").Substring(0, 12);
            string marker = "OBSERVED_" + Guid.NewGuid().ToString("N");
            var json = new JavaScriptSerializer();
            var report = new Dictionary<string, object> {
                ["State"] = "RUNNING", ["Model"] = model, ["Endpoint"] = endpoint,
                ["Project"] = projectName, ["Module"] = moduleName,
                ["MarkerSha256"] = EditorDocument.Hash(marker), ["StartedUtc"] = DateTime.UtcNow.ToString("o"),
                ["AssemblyPath"] = typeof(LlmChatClient).Assembly.Location,
                ["AssemblyModuleVersionId"] = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                ["ToolInvocations"] = 0, ["MacroExecutionRequested"] = false,
                ["Scope"] = "Production HTTP/client/tools/session and native Excel COM. In-process test dispatch; installed bridge and embedded UI are not qualified."
            };
            object application = null, books = null, book = null, vbe = null, project = null, components = null, component = null, code = null;
            string before = null;
            ExceptionDispatchInfo failure = null;
            var watch = Stopwatch.StartNew();
            try
            {
                var provider = LlmProvider.All.Single(item => item.IsOllama);
                var settings = new LlmSettings { ProviderName = provider.Name, OllamaEndpoint = endpoint,
                    OllamaModel = model, VbeEditApproval = "ReadOnly" };
                // Shadow any inherited credential without loading or saving personal settings.
                settings.SetKey(provider, "vbai-synthetic-loopback-excel-test");
                var models = AwaitOnSta(LlmChatClient.ListModelsAsync(provider, settings), 125);
                Assert.IsTrue(models.Any(item => item.Id == model), "The local model must already be installed; this test never downloads it.");

                application = Activator.CreateInstance(excelType);
                lifetime.Capture(application);
                dynamic excel = application;
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                books = excel.Workbooks;
                book = ((dynamic)books).Add();
                project = ((dynamic)book).VBProject;
                ((dynamic)project).Name = projectName;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Add(1);
                ((dynamic)component).Name = moduleName;
                code = ((dynamic)component).CodeModule;
                ((dynamic)code).AddFromString("Option Explicit\r\nPublic Const ObservedMarker As String = \"" + marker + "\"");
                before = ReadCode(code);
                report["BeforeSha256"] = EditorDocument.Hash(before);
                vbe = excel.VBE;
                var adapter = new EditorVbeModule(vbe, project, component);
                report["HostProcessId"] = adapter.HostProcessId;
                var tools = new LlmVbeTools(new VbeSession(vbe), null, settings) {
                    BoundProject = projectName, CurrentProviderName = provider.Name, Mode = ChatMode.Discussion
                };
                // Retain the production definition; restrict only the advertised tool set.
                object[] definitions = LlmVbeTools.Definitions.Where(definition => {
                    var item = (IDictionary<string, object>)json.DeserializeObject(json.Serialize(definition));
                    return (string)((IDictionary<string, object>)item["function"])["name"] == "read_module";
                }).ToArray();
                Assert.AreEqual(1, definitions.Length);
                var messages = new List<object> {
                    new { role = "system", content = "You inspect a disposable VBA project in read-only Discussion mode. Call read_module exactly once with the Project and Module provided by the user. Never invent source or execute macros. After its result, return the exact string constant named ObservedMarker." },
                    new { role = "user", content = "Read Project=" + projectName + ", Module=" + moduleName + " using read_module. What is the exact value of ObservedMarker? Its value is only in the native module." }
                };
                Assert.IsFalse(json.Serialize(messages).Contains(marker), "The answer must not be supplied in the prompt.");
                Assert.IsFalse(json.Serialize(definitions).Contains(marker));
                using (var client = new LlmChatClient(provider, settings, model))
                {
                    int fragments = 0;
                    client.TextDelta = text => { if (!string.IsNullOrEmpty(text)) Interlocked.Increment(ref fragments); };
                    var requested = AwaitOnSta(client.CompleteAsync(messages, definitions), 125);
                    Assert.IsTrue(requested.ContainsKey("tool_calls"), "The model must request the native read tool.");
                    var calls = (object[])requested["tool_calls"];
                    Assert.AreEqual(1, calls.Length, "No extra or repeated tool calls are dispatched.");
                    var call = (IDictionary<string, object>)calls[0];
                    var function = (IDictionary<string, object>)call["function"];
                    Assert.AreEqual("read_module", function["name"]);
                    string arguments = (string)function["arguments"];
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    Assert.IsNotNull(values);
                    Assert.AreEqual(projectName, values["Project"]);
                    Assert.AreEqual(moduleName, values["Module"]);
                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId, "Native tool dispatch must stay on the owning STA.");
                    report["ToolInvocations"] = 1;
                    string result = tools.Invoke("read_module", arguments);
                    var reply = json.Deserialize<Response>(result);
                    Assert.IsTrue(reply.Ok, "Production native read_module failed: " + reply.Error);
                    var data = (IDictionary<string, object>)json.DeserializeObject(json.Serialize(reply.Data));
                    StringAssert.Contains((string)data["Code"], marker);
                    Assert.AreEqual(EditorDocument.Normalize(before), EditorDocument.Normalize((string)data["Code"]),
                        "The tool result must contain the independently read native module.");
                    report["ToolCodeSha256"] = data["Sha256"];
                    messages.Add(requested);
                    messages.Add(new { role = "tool", tool_call_id = (string)call["id"], content = result });
                    messages.Add(new { role = "user", content = "Using that tool result, reply with only the exact ObservedMarker string value. Do not call any more tools." });
                    var answer = AwaitOnSta(client.CompleteAsync(messages, new object[0]), 125);
                    Assert.IsFalse(answer.ContainsKey("tool_calls") && ((object[])answer["tool_calls"]).Length != 0);
                    string final = Convert.ToString(answer["content"]);
                    StringAssert.Contains(final, marker, "The final model answer must contain the unpredictable value obtained from native VBA.");
                    report["FinalContainsNativeMarker"] = true;
                    report["FinalAnswerSha256"] = EditorDocument.Hash(final);
                    report["StreamFragments"] = fragments;
                }
                report["State"] = "PASS";
            }
            catch (Exception error) { failure = ExceptionDispatchInfo.Capture(error); report["State"] = "FAIL"; report["Error"] = error.ToString(); }
            finally
            {
                try
                {
                    if (before != null)
                    {
                        Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                        string after = ReadCode(code);
                        report["AfterSha256"] = EditorDocument.Hash(after);
                        report["CodeUnchanged"] = before == after;
                        Assert.AreEqual(before, after, "The entire native source must remain byte-for-byte unchanged during the model/tool conversation.");
                        Assert.AreEqual(2, (int)((dynamic)project).Mode);
                        Assert.AreEqual("", (string)((dynamic)book).Path, "The disposable workbook must remain unsaved.");
                    }
                }
                catch (Exception error)
                {
                    report["State"] = "FAIL"; report["ReadbackError"] = error.ToString();
                    if (failure == null) failure = ExceptionDispatchInfo.Capture(error);
                }
                Release(code); Release(component); Release(components); Release(project); Release(vbe);
                try { if (book != null) ((dynamic)book).Close(false); }
                catch (Exception error) { report["WorkbookCleanupError"] = error.Message; }
                Release(book); Release(books);
                try { if (application != null && lifetime.OwnsApplication) ((dynamic)application).Quit(); }
                catch (Exception error) { report["ApplicationCleanupError"] = error.Message; }
                Release(application);
                report["ElapsedMilliseconds"] = watch.ElapsedMilliseconds;
                report["CompletedUtc"] = DateTime.UtcNow.ToString("o");
                string path = Path.Combine(context.TestRunResultsDirectory, "ollama-excel-" + projectName + ".json");
                Directory.CreateDirectory(context.TestRunResultsDirectory);
                File.WriteAllText(path, json.Serialize(report));
                context.AddResultFile(path);
                context.WriteLine(json.Serialize(report));
            }
            failure?.Throw();
        }

        private static string ReadCode(object code)
        {
            int count = (int)((dynamic)code).CountOfLines;
            return count == 0 ? "" : (string)((dynamic)code).Lines[1, count];
        }

        private static T AwaitOnSta<T>(Task<T> task, int seconds)
        {
            var watch = Stopwatch.StartNew();
            while (!task.IsCompleted && watch.Elapsed.TotalSeconds < seconds)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(10);
            }
            Assert.IsTrue(task.IsCompleted, "Live Ollama request exceeded the bounded wait; no request or native operation will be retried.");
            return task.GetAwaiter().GetResult();
        }

        private static void Release(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.FinalReleaseComObject(value); }
            catch (COMException error) { Console.WriteLine("OllamaExcel COM cleanup: " + error.Message); }
        }
    }
}
