using System;
using System.Collections.Generic;
using System.Reflection;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        public sealed class LiveCutSnapshot { public object[] Selected { get; set; } = new object[] { "A", "B" }; }
        public sealed class LiveCutResult
        {
            public string DesignerClipboardRecoveryId { get; set; }
            public bool DesignerChangeObserved { get; set; }
            public LiveCutSnapshot Before { get; set; } = new LiveCutSnapshot();
        }
        public sealed class NativeHistoryResult { public object[] Changes { get; set; } }
        public sealed class LiveSourceResult { public string Code { get; set; } = "abc"; public string Sha256 { get; set; } = "sha"; }

        [TestMethod]
        public void NativeCutNotificationsRequireBothRecoveryIdentityAndObservedDesignerChange()
        {
            foreach (string identity in new[] { null, "", "live" })
            foreach (bool observed in new[] { false, true })
            foreach (bool subscribed in new[] { false, true })
            {
                var tools = new ToolFixture().Tools; int notifications = 0;
                if (subscribed) tools.FormCut += change => {
                    notifications++; Assert.AreSame(tools, change.Owner); Assert.AreEqual(2, change.ControlCount);
                    Assert.AreEqual("live", change.RecoveryId); Assert.AreEqual("P", change.Project);
                };
                tools.Execute = r => Response.Success(new LiveCutResult { DesignerClipboardRecoveryId = identity, DesignerChangeObserved = observed });
                var args = Arguments("native_form_clipboard"); args["Action"] = "cut";
                Success(tools.Invoke("native_form_clipboard", Json.Serialize(args)), "native cut");
                Assert.AreEqual(identity == "live" && observed && subscribed ? 1 : 0, notifications);
            }
            foreach (bool restoring in new[] { false, true })
            foreach (bool subscribed in new[] { false, true })
            {
                var tools = new ToolFixture().Tools; int notifications = 0;
                if (subscribed) tools.CodeEdited += change => notifications++;
                typeof(LlmVbeTools).GetField("restoring", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(tools, restoring);
                tools.Execute = r => Response.Success(new NativeHistoryResult { Changes = new[] { new CodeChange(), new CodeChange() } });
                Success(tools.Invoke("native_code_history", Json.Serialize(Arguments("native_code_history"))), "native history");
                Assert.AreEqual(!restoring && subscribed ? 2 : 0, notifications);
            }
        }

        [TestMethod]
        public void FailedNativeCutAndPasteStillReadBackPartialMutations()
        {
            foreach (string command in new[] { "cut_code", "paste_code" })
            {
                var tools = new ToolFixture().Tools; int reads = 0;
                tools.Execute = r => {
                    if (r.Command == "read_module") { reads++; return Response.Success(new LiveSourceResult { Code = reads == 1 ? "abc" : "partial" }); }
                    throw new InvalidOperationException("native write interrupted");
                };
                var args = Arguments(command); args["ExpectedSha256"] = "sha";
                var response = Json.Deserialize<Response>(tools.Invoke(command, Json.Serialize(args)));
                Assert.IsFalse(response.Ok); StringAssert.Contains(response.Error, "native write interrupted");
                Assert.AreEqual(2, reads);
            }
        }

        [TestMethod]
        public void ListInitializerArgumentsRequireBoundedPrintableRectangularRows()
        {
            var tools = new ToolFixture().Tools;
            tools.Execute = r => r.Command == "read_module" ? Response.Success(new LiveSourceResult()) : Response.Success(new object());
            var invalid = new object[] { null, 12, new object[] { 12 }, new object[] { new string[0] },
                new object[] { new string[11] }, new object[] { new object[] { 12 } }, new object[] { new object[] { null } },
                new object[] { new[] { new string('x', 257) } }, new object[] { new[] { "bad\ncell" } },
                new object[] { new[] { "a" }, new[] { "a", "b" } }, new object[65] };
            foreach (object rows in invalid)
            {
                var args = Arguments("set_form_list_initializer"); args.Remove("Items"); args["Rows"] = rows;
                args["ExpectedSha256"] = "sha";
                Failed(tools.Invoke("set_form_list_initializer", Json.Serialize(args)), "invalid matrix");
            }
            foreach (object rows in new object[] { new object[0], new[] { new[] { "" } }, new[] { new[] { "a", "b" }, new[] { "c", "d" } } })
            {
                var args = Arguments("set_form_list_initializer"); args.Remove("Items"); args["Rows"] = rows;
                args["ExpectedSha256"] = "sha";
                Success(tools.Invoke("set_form_list_initializer", Json.Serialize(args)), "valid matrix");
            }
        }
    }
}
