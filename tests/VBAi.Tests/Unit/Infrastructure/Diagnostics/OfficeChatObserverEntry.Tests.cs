using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Unit
{
    /// <summary>The observer protocol refuses replacements and unknown results before Word teardown.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeChatObserverEntryTests
    {
        private static OfficeChatObserverPlan Plan() => new OfficeChatObserverPlan
        {
            Version = 1, Nonce = Guid.NewGuid().ToString("N"), UserSid = "S-1-5-21-test",
            Helper = @"C:\candidate\VBAi.Desktop.Helper.exe", HelperSha256 = "helper",
            Tests = @"C:\candidate\VBAi.Tests.dll", TestsSha256 = "tests", TestsMvid = Guid.NewGuid().ToString("D"),
            Product = @"C:\candidate\VBAi.dll", ProductSha256 = "product", ProductMvid = Guid.NewGuid().ToString("D"),
            ParentPid = 42, ParentBirthUtc = new DateTime(2026, 10, 9, 4, 0, 0, DateTimeKind.Utc).ToString("o"),
            ParentImage = @"C:\runner\testhost.exe", HostPid = 51,
            HostBirthUtc = new DateTime(2026, 10, 9, 4, 1, 0, DateTimeKind.Utc).ToString("o"),
            HostImage = @"C:\Office\WINWORD.EXE", Label = "Disposable", Project = @"C:\owned\Disposable.docm",
            Module = "Q028Marker", Marker = "NATIVE_" + Guid.NewGuid().ToString("N"),
            Model = "qwen3.5:9b", Profile = "frozen-profile", Culture = "fr-FR",
            Receipt = @"C:\owned\observer-receipt.json"
        };

        private static void Validate(OfficeChatObserverPlan plan, string changed = null) =>
            OfficeChatObserverEntry.Validate(plan, @"C:\owned\observer-plan.json",
                changed == "ownImage" ? @"C:\other\VBAi.Desktop.Helper.exe" : @"C:\candidate\VBAi.Desktop.Helper.exe",
                path => path.EndsWith("observer-receipt.json") ? changed == "existingReceipt" : changed != "missingAssembly",
                path => changed == "hash" ? "replacement" : path.EndsWith("VBAi.Desktop.Helper.exe") ? "helper" :
                    path.EndsWith("VBAi.Tests.dll") ? "tests" : "product");

        [TestMethod] public void ExactOwnedCandidateAndFreshReceiptAreAdmitted() => Validate(Plan());

        [DataTestMethod]
        [DataRow("version")][DataRow("nonce")][DataRow("testsMvid")][DataRow("productMvid")]
        [DataRow("pid")][DataRow("samePid")][DataRow("sid")][DataRow("label")][DataRow("project")]
        [DataRow("module")][DataRow("marker")][DataRow("markerGuid")][DataRow("model")][DataRow("profile")]
        [DataRow("parentBirth")][DataRow("hostBirth")][DataRow("localBirth")][DataRow("relative")][DataRow("network")]
        [DataRow("helperName")][DataRow("testsName")][DataRow("productName")][DataRow("hostImage")]
        [DataRow("testsFolder")][DataRow("productFolder")][DataRow("receiptFolder")][DataRow("receiptName")]
        [DataRow("ownImage")][DataRow("existingReceipt")][DataRow("missingAssembly")][DataRow("hash")]
        public void InvalidManifestNeverReachesUiAutomation(string change)
        {
            var p = Plan();
            switch (change)
            {
                case "version": p.Version = 2; break;
                case "nonce": p.Nonce = "bad"; break;
                case "testsMvid": p.TestsMvid = "bad"; break;
                case "productMvid": p.ProductMvid = "bad"; break;
                case "pid": p.HostPid = 0; break;
                case "samePid": p.HostPid = p.ParentPid; break;
                case "sid": p.UserSid = null; break;
                case "label": p.Label = " "; break;
                case "project": p.Project = null; break;
                case "module": p.Module = "Production"; break;
                case "marker": p.Marker = "sample"; break;
                case "markerGuid": p.Marker = "NATIVE_" + new string('z', 32); break;
                case "model": p.Model = " "; break;
                case "profile": p.Profile = null; break;
                case "parentBirth": p.ParentBirthUtc = "bad"; break;
                case "hostBirth": p.HostBirthUtc = "bad"; break;
                case "localBirth": p.HostBirthUtc = "2026-10-09T04:01:00.0000000"; break;
                case "relative": p.Tests = "VBAi.Tests.dll"; break;
                case "network": p.Tests = @"\\server\candidate\VBAi.Tests.dll"; break;
                case "helperName": p.Helper = @"C:\candidate\other.exe"; break;
                case "testsName": p.Tests = @"C:\candidate\other.dll"; break;
                case "productName": p.Product = @"C:\candidate\other.dll"; break;
                case "hostImage": p.HostImage = @"C:\Office\EXCEL.EXE"; break;
                case "testsFolder": p.Tests = @"C:\replacement\VBAi.Tests.dll"; break;
                case "productFolder": p.Product = @"C:\replacement\VBAi.dll"; break;
                case "receiptFolder": p.Receipt = @"C:\other\observer-receipt.json"; break;
                case "receiptName": p.Receipt = @"C:\owned\other.json"; break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => Validate(p, change));
        }

        private static OfficeChatObserverReceipt Receipt(OfficeChatObserverPlan p) => new OfficeChatObserverReceipt
        {
            Nonce = p.Nonce, HostPid = p.HostPid, HostBirthUtc = p.HostBirthUtc, ProcessId = 91,
            TestsMvid = p.TestsMvid, ProductMvid = p.ProductMvid, Completed = true, Settled = true,
            StreamedWhileBusy = true, CancellationAcknowledged = true, NextSendCompleted = true, NativeMarkerVisible = true
        };

        [TestMethod]
        public void SettledFunctionalPassRequiresOriginalObserverExit()
        {
            var p = Plan(); OfficeChatObserverEntry.RequireTerminal(p, Receipt(p), 91, true, 0);
        }

        [TestMethod]
        public void KnownSettledOracleFailureMayCloseHostButDoesNotPassQualification()
        {
            var p = Plan(); var r = Receipt(p); r.NativeMarkerVisible = false; r.Error = "Marker oracle failed";
            OfficeChatObserverEntry.RequireTerminal(p, r, 91, true, 1);
        }

        [DataTestMethod]
        [DataRow("null")][DataRow("nonce")][DataRow("child")][DataRow("host")][DataRow("birth")]
        [DataRow("tests")][DataRow("product")][DataRow("incomplete")][DataRow("pending")][DataRow("noExit")]
        [DataRow("crash")][DataRow("failedExit")][DataRow("missingStream")][DataRow("missingCancel")]
        [DataRow("missingNext")][DataRow("missingMarker")][DataRow("errorOnSuccess")][DataRow("silentFailure")]
        public void UnknownOrContradictoryTerminalRetainsHostWithoutReplaying(string change)
        {
            var p = Plan(); var r = Receipt(p); uint code = 0;
            switch (change)
            {
                case "null": r = null; break;
                case "nonce": r.Nonce = Guid.NewGuid().ToString("N"); break;
                case "child": r.ProcessId++; break;
                case "host": r.HostPid++; break;
                case "birth": r.HostBirthUtc = DateTime.UtcNow.ToString("o"); break;
                case "tests": r.TestsMvid = Guid.NewGuid().ToString("D"); break;
                case "product": r.ProductMvid = Guid.NewGuid().ToString("D"); break;
                case "incomplete": r.Completed = false; break;
                case "pending": r.Settled = false; break;
                case "crash": code = 3; break;
                case "failedExit": code = 1; break;
                case "missingStream": r.StreamedWhileBusy = false; break;
                case "missingCancel": r.CancellationAcknowledged = false; break;
                case "missingNext": r.NextSendCompleted = false; break;
                case "missingMarker": r.NativeMarkerVisible = false; break;
                case "errorOnSuccess": r.Error = "Error"; break;
                case "silentFailure": r.NativeMarkerVisible = false; code = 1; break;
            }
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeChatObserverEntry.RequireTerminal(p, r, 91, change != "noExit", code));
        }
    }
}
