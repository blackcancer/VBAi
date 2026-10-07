namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using VBAi;

    [TestClass]
    public sealed partial class VbeSolidWorksMacroPublicationTests
    {
        [DataTestMethod, TestCategory("Unit")]
        [DataRow("type"), DataRow("protected"), DataRow("help"), DataRow("context"), DataRow("conditional"),
         DataRow("document"), DataRow("reserved"), DataRow("broken-reference"), DataRow("project-reference"), DataRow("collision"), DataRow("version")]
        public async Task PublicationPreflightRefusesUnsupportedSourceBeforeAnyExportOrNativeEntry(string fault)
        {
            using (var f = new Fixture())
            {
                if (fault == "type") f.Source.Type = 100;
                if (fault == "protected") f.Source.Protection = 1;
                if (fault == "help") f.Source.HelpFile = @"C:\private\help.chm";
                if (fault == "context") f.Source.HelpContextID = 8;
                if (fault == "conditional") f.ConditionalCompilation = "TRACE=1";
                if (fault == "document") f.Source.VBComponents.Items[0].Type = 100;
                if (fault == "reserved") f.Source.VBComponents.Items[0].Name = "ThisLibrary";
                if (fault == "broken-reference") f.Source.References.Items.Add(new Reference { IsBroken = true });
                if (fault == "project-reference") f.Source.References.Items.Add(new Reference { Type = 1 });
                if (fault == "collision") f.Path = System.IO.Path.Combine(f.Root, "Source.swp");
                var request = f.Request();
                if (fault == "version") request.ExpectedProjectVersion = "stale";
                try
                {
                    var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(request);
                    Assert.IsFalse(result.Verified, result.Error);
                    Assert.IsTrue(result.Terminal, result.Error);
                    Assert.IsFalse(result.MutationInvoked, result.Error);
                }
                catch (InvalidOperationException) { }
                Assert.AreEqual(0, f.CreateAttempts);
                Assert.AreEqual(0, f.Source.VBComponents.Items.Sum(c => c.ExportAttempts));
                Assert.AreEqual(0, Directory.GetDirectories(f.Root).Length);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationRequiresNativeGeneralReaderWithoutInventedComProperty()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(typeof(Project).GetProperty("ConditionalCompilationArguments"));
                f.Service.PublicationReadGeneral = null;
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Service.PublishSolidWorksMacroAsync(f.Request()));
                Assert.AreEqual(0, f.CreateAttempts);
                Assert.AreEqual(0, f.Source.VBComponents.Items[0].ExportAttempts);
                Assert.IsFalse(File.Exists(f.Path));
            }
        }

        [DataTestMethod, TestCategory("Unit")]
        [DataRow("Available"), DataRow("Terminal"), DataRow("OriginalExecuteReturned"), DataRow("DialogClosed"),
         DataRow("Uncertain"), DataRow("Error"), DataRow("CancelAttempts"), DataRow("FieldAttempts"),
         DataRow("OkAttempts"), DataRow("MutationInvoked"), DataRow("CommittedRequested"), DataRow("OpenAttempts"),
         DataRow("OptionsVersion"), DataRow("ConditionalCompilation")]
        public async Task PublicationRefusesIncompleteOrMutatingGeneralProofBeforeExports(string fault)
        {
            using (var f = new Fixture())
            {
                var proof = f.GeneralSnapshot();
                if (fault == "Error") proof[fault] = "Native read failed";
                else if (fault == "OptionsVersion") proof.Remove(fault);
                else if (fault == "ConditionalCompilation") proof[fault] = "TRACE=1";
                else if (proof[fault] is bool) proof[fault] = !(bool)proof[fault];
                else proof[fault] = (int)proof[fault] + 1;
                f.Service.PublicationReadGeneral = request => Task.FromResult<object>(proof);
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified, result.Error);
                Assert.AreEqual(0, f.CreateAttempts, result.Error);
                Assert.AreEqual(0, f.Source.VBComponents.Items[0].ExportAttempts, result.Error);
                Assert.IsFalse(File.Exists(f.Path));
                Assert.AreEqual(0, Directory.GetDirectories(f.Root).Length);
                Assert.IsFalse(result.RetryAllowed);
            }
        }

        [DataTestMethod, TestCategory("Unit"), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5)]
        public async Task PublicationRechecksCanonicalSourceAcrossEveryGeneralAwait(int readOrdinal)
        {
            using (var f = new Fixture())
            {
                f.AfterGeneralRead = ordinal => { if (ordinal == readOrdinal) f.Source.VBComponents.Items[0].CodeModule.Source = "Changed during modal"; };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified, result.Error);
                Assert.IsFalse(result.OriginalPreserved, result.Error);
                Assert.AreEqual(readOrdinal <= 2 ? 0 : 1, f.CreateAttempts, result.Error);
                Assert.AreEqual(readOrdinal == 5 ? 1 : 0, f.SaveAttempts, result.Error);
                Assert.AreEqual(readOrdinal >= 3, result.DestinationCreated, result.Error);
                Assert.IsFalse(result.RetryAllowed);
            }
        }

        [DataTestMethod, TestCategory("Unit"), DataRow(2), DataRow(3), DataRow(4), DataRow(5)]
        public async Task PublicationRejectsNativeOptionsDriftWithAccurateCreatedPartial(int readOrdinal)
        {
            using (var f = new Fixture())
            {
                f.BeforeGeneralRead = ordinal => { if (ordinal == readOrdinal) f.GeneralOptionsVersion = "changed-options"; };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified, result.Error);
                Assert.IsTrue(result.Terminal, result.Error);
                Assert.AreEqual(readOrdinal == 2 ? 0 : 1, f.CreateAttempts, result.Error);
                Assert.AreEqual(readOrdinal == 5 ? 1 : 0, f.SaveAttempts, result.Error);
                Assert.AreEqual(readOrdinal > 2, File.Exists(f.Path), result.Error);
                Assert.IsFalse(result.RollbackPerformed);
                Assert.IsFalse(result.RetryAllowed);
            }
        }

        [DataTestMethod, TestCategory("Unit"), DataRow(1), DataRow(3)]
        public async Task PublicationRefusesReplacedCanonicalSourceAfterGeneralAwait(int readOrdinal)
        {
            using (var f = new Fixture())
            {
                f.AfterGeneralRead = ordinal =>
                {
                    if (ordinal == readOrdinal) f.Vbe.VBProjects[0] = new Project { Name = f.Source.Name };
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified, result.Error);
                Assert.IsFalse(result.OriginalPreserved, result.Error);
                Assert.AreEqual(readOrdinal == 1 ? 0 : 1, f.CreateAttempts, result.Error);
                Assert.AreEqual(0, f.SaveAttempts, result.Error);
                Assert.IsFalse(result.RetryAllowed);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationCopiesHiddenClassAttributesAndRetainsOriginalIdentityAndExports()
        {
            using (var f = new Fixture())
            {
                f.Source.VBComponents.Items.Add(new Component("Class1", 2));
                var original = f.Source;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsTrue(result.Verified, result.Error); Assert.IsTrue(result.Terminal, result.Error);
                Assert.IsFalse(result.Uncertain); Assert.IsFalse(result.RetryAllowed);
                Assert.IsTrue(result.OriginalPreserved); Assert.IsTrue(result.IdentityChanged);
                Assert.AreSame(original, f.Vbe.VBProjects[0]);
                Assert.AreEqual("Source", result.OriginalProject);
                Assert.IsNull(result.OriginalHostPath, "The real pathless Type101 getter returns an empty string, not a file to hash.");
                Assert.AreEqual("", f.Source.FileName, "Normalization must not write to the original source getter.");
                Assert.AreEqual("Published", result.DestinationProject);
                Assert.AreEqual(1, f.CreateAttempts); Assert.AreEqual(1, f.SaveAttempts);
                Assert.AreEqual(5, f.GeneralReads, "Initial, pre-create, post-create, pre-save and after-save native inspections are mandatory.");
                Assert.AreEqual(2, f.Target.VBComponents.ImportAttempts);
                Assert.AreEqual(1, f.Target.VBComponents.RemoveAttempts);
                Assert.AreEqual(3, f.Target.VBComponents.Items.Count);
                Assert.AreEqual(3, Directory.GetFiles(result.StagingPath, "*.cls").Length, "Source, imported verification and distinct verified-save export are retained.");
                Assert.AreEqual("Original description", f.Target.Description);
                Assert.IsTrue(result.Claims.Any(c => c.Phase == "BeforeSave"));
                var copy = result.Claims; copy[0] = null;
                Assert.IsNotNull(result.Claims[0], "Receipt entries cannot be replaced by the caller.");
            }
        }

        [DataTestMethod, TestCategory("Unit"), DataRow(false), DataRow(true)]
        public async Task PublicationTransportPreservesFormResourcePairAndRejectsDesignerMismatch(bool mismatch)
        {
            using (var f = new Fixture())
            {
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                f.Service.PublicationFormTree = (project, form) => new Dictionary<string, object>
                {
                    ["TreeVersion"] =
                    mismatch && project == "Published" ? "changed-picture-or-control" : "full-controls-picture-digest",
                    ["Properties"] = new[] { new { Name = "Caption", Value = mismatch && project == "Published" ? "Different" : "Original", Error = (string)null } },
                    ["Controls"] = new object[0]
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.AreEqual(!mismatch, result.Verified, result.Error);
                Assert.AreEqual(mismatch ? 0 : 1, f.SaveAttempts);
                Assert.IsTrue(File.Exists(System.IO.Path.Combine(result.StagingPath, "Form1.frx")), result.Error);
                Assert.AreEqual(1, f.CreateAttempts);
                Assert.IsFalse(result.RetryAllowed);
                if (!mismatch) Assert.IsTrue(File.Exists(System.IO.Path.Combine(result.StagingPath, "verified-Form1.frx")));
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationPreservesFreshNativeReferencesAndAddsExactSourceVersions()
        {
            using (var f = new Fixture())
            {
                f.Source.References.Items.Add(new Reference());
                f.NativeReference = new Reference { GUID = "{83A33D22-27C5-11CE-BFD4-00400513BB57}" };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsTrue(result.Verified, result.Error);
                Assert.AreEqual(2, f.Target.References.Items.Count);
                Assert.IsNotNull(result.AddedNativeHostReferences);
                Assert.AreEqual(1, f.Target.References.AddAttempts);
            }
        }

        [DataTestMethod, TestCategory("Unit"), DataRow(false), DataRow(true)]
        public async Task PublicationImportFaultReturnsHonestRetainedPartialWithoutRetryOrSave(bool throwsAfterAdd)
        {
            using (var f = new Fixture())
            {
                f.ImportFault = throwsAfterAdd ? "throw" : "wrong-code";
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Uncertain); Assert.IsTrue(result.Terminal);
                Assert.IsFalse(result.RetryAllowed); Assert.IsFalse(result.RollbackPerformed);
                Assert.IsNotNull(f.Target, result.Error);
                Assert.AreEqual(1, f.Target.VBComponents.ImportAttempts); Assert.AreEqual(0, f.SaveAttempts);
                Assert.AreEqual(2, f.Vbe.VBProjects.Count);
                Assert.IsTrue(File.Exists(f.Path)); Assert.IsTrue(Directory.Exists(result.StagingPath));
                Assert.AreEqual("Option Explicit", f.Source.VBComponents.Items[0].CodeModule.Source);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationUnsettledCreationDoesNotReadDestinationOrImport()
        {
            using (var f = new Fixture())
            {
                f.CreationUnsettled = true;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Terminal, result.Error); Assert.IsTrue(result.Uncertain, result.Error);
                Assert.IsFalse(result.Verified); Assert.AreEqual(0, f.Target.VBComponents.ImportAttempts);
                Assert.AreEqual(0, f.SaveAttempts); Assert.IsFalse(result.OriginalPreserved);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationRejectsSourceChangeDuringNativeCreationBeforeFirstImport()
        {
            using (var f = new Fixture())
            {
                f.AfterCreate = () => f.Source.VBComponents.Items[0].CodeModule.Source = "Changed";
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Uncertain);
                Assert.IsFalse(result.OriginalPreserved, result.Error); Assert.IsNotNull(f.Target, result.Error);
                Assert.AreEqual(0, f.Target.VBComponents.ImportAttempts);
                Assert.AreEqual(0, f.SaveAttempts);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationCreationCachedAuthorizationDoesNotReadComButStillRejectsRevokedPolicy()
        {
            using (var f = new Fixture())
            {
                bool revoked = false;
                int falseChecks = 0;
                f.Service.PublicationCreate = (request, authorize, claim, context) =>
                {
                    f.Source.VBComponents.Items[0].CodeModule.FailRead = true;
                    authorize(false); // A pure final check must succeed despite inaccessible COM source getters.
                    falseChecks++;
                    revoked = true;
                    Assert.ThrowsException<InvalidOperationException>(() => authorize(false));
                    f.Source.VBComponents.Items[0].CodeModule.FailRead = false;
                    return Task.FromResult(new VbeProjectComponents.SolidWorksMacroCreationResult
                    {
                        Terminal = true,
                        Verified = false,
                        Error = "Cached policy refused before native entry."
                    });
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request(), authorization: shared =>
                    {
                        if (revoked) throw new InvalidOperationException("Policy revoked.");
                    });
                Assert.AreEqual(1, falseChecks, result.Error);
                Assert.AreEqual(0, f.CreateAttempts);
                Assert.AreEqual(0, f.SaveAttempts);
                Assert.IsTrue(result.Terminal);
                Assert.IsFalse(result.Verified);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public async Task PublicationLostContextBeforeExportDoesNotInvokeNativeCreation()
        {
            using (var f = new Fixture())
            {
                int entries = 0;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)
                    await f.Service.PublishSolidWorksMacroAsync(f.Request(), requireNativeContext: () =>
                    {
                        if (++entries > 3) throw new InvalidOperationException("Desktop ownership changed.");
                    });
                Assert.IsFalse(result.Verified); Assert.AreEqual(0, f.CreateAttempts);
                Assert.AreEqual(0, f.Source.VBComponents.Items[0].ExportAttempts);
            }
        }

        [TestMethod, TestCategory("Unit")]
        public void PublicationRequiresReadableCompleteDesignerAndComparesHiddenAttributes()
        {
            var nested = new Dictionary<string, object>
            {
                ["Controls"] = new object[] {
                new Dictionary<string, object> { ["Properties"] = new object[] {
                    new Dictionary<string, object> { ["Name"] = "Picture", ["Error"] = "Cannot fingerprint image" } } } }
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectComponents.RequirePublicationDesignerReadable(nested));
            Assert.AreEqual(VbeProjectComponents.PublicationAttributes("Attribute VB_Name = \"Class1\"\r\nAttribute VB_PredeclaredId = True\r\n"),
                VbeProjectComponents.PublicationAttributes("Attribute VB_PredeclaredId = True\nAttribute VB_Name = \"Class1\"\n"));
            Assert.AreNotEqual(VbeProjectComponents.PublicationAttributes("Attribute VB_PredeclaredId = True"),
                VbeProjectComponents.PublicationAttributes("Attribute VB_PredeclaredId = False"));
        }

        // Insert inside the existing matching VbeSolidWorksMacroPublicationTests mirror.
        // Native shapes below are synthetic Q020 source wire0062; tests never touch a host.
        private static Dictionary<string, object> CapturedNativeFontTree() =>
            new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Dictionary<string, object>>(@"{
  ""TreeVersion"": ""captured-managed-and-ole-font-shapes"",
  ""Properties"": [
    {
      ""Name"": ""_Font_Reserved"",
      ""Type"": ""System.Drawing.Font"",
      ""Kind"": ""scalar"",
      ""ReadOnly"": false,
      ""AllowedValues"": null,
      ""SetterStatus"": ""GetterUnavailable"",
      ""Value"": null,
      ""Display"": null,
      ""Digest"": null,
      ""Error"": ""Propriété ou méthode non gérée par cet objet"",
      ""NumIndices"": 0,
      ""Members"": null
    },
    {
      ""Name"": ""Font"",
      ""Type"": ""System.Drawing.Font"",
      ""Kind"": ""object"",
      ""ReadOnly"": false,
      ""AllowedValues"": null,
      ""SetterStatus"": ""DescriptorCandidateUnverified"",
      ""Value"": null,
      ""Display"": ""Font"",
      ""Digest"": null,
      ""Error"": null,
      ""NumIndices"": 0,
      ""Members"": [
        {
          ""Name"": ""FontFamily"",
          ""Type"": ""System.Drawing.FontFamily"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": ""[FontFamily: Name=Tahoma]"",
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Bold"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""GdiCharSet"",
          ""Type"": ""System.Byte"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": 1,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""GdiVerticalFont"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Italic"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Name"",
          ""Type"": ""System.String"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": ""Tahoma"",
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""OriginalFontName"",
          ""Type"": ""System.String"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": null,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Strikeout"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Underline"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Style"",
          ""Type"": ""System.Drawing.FontStyle"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": [
            ""Regular"",
            ""Bold"",
            ""Italic"",
            ""Underline"",
            ""Strikeout""
          ],
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": ""Regular"",
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Size"",
          ""Type"": ""System.Single"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": 8.25,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""SizeInPoints"",
          ""Type"": ""System.Single"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": 8.25,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Unit"",
          ""Type"": ""System.Drawing.GraphicsUnit"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": [
            ""World"",
            ""Display"",
            ""Pixel"",
            ""Point"",
            ""Inch"",
            ""Document"",
            ""Millimeter""
          ],
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": ""Point"",
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Height"",
          ""Type"": ""System.Int32"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": 14,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""IsSystemFont"",
          ""Type"": ""System.Boolean"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": false,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""SystemFontName"",
          ""Type"": ""System.String"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": true,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorReadOnly"",
          ""Value"": """",
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        }
      ]
    },
    {
      ""Name"": ""Picture"",
      ""Type"": ""stdole.IPictureDisp"",
      ""Kind"": ""object"",
      ""ReadOnly"": false,
      ""AllowedValues"": null,
      ""SetterStatus"": ""DescriptorCandidateUnverified"",
      ""Value"": null,
      ""Display"": ""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",
      ""Digest"": ""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",
      ""Error"": null,
      ""NumIndices"": 0,
      ""Members"": null
    }
  ],
  ""Controls"": [
    {
      ""Name"": ""Q020Label"",
      ""Type"": ""Label"",
      ""Properties"": [
        {
          ""Name"": ""_Font_Reserved"",
          ""Type"": ""System.Drawing.Font"",
          ""Kind"": null,
          ""ReadOnly"": false,
          ""AllowedValues"": null,
          ""SetterStatus"": ""GetterUnavailable"",
          ""Value"": null,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": ""L'opération GetValue sur le composant a échoué avec le code d'erreur 0x80020003."",
          ""NumIndices"": 0,
          ""Members"": null
        },
        {
          ""Name"": ""Font"",
          ""Type"": ""System.Drawing.Font"",
          ""Kind"": ""object"",
          ""ReadOnly"": false,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorCandidateUnverified"",
          ""Value"": null,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": [
            {
              ""Name"": ""Name"",
              ""Type"": ""System.String"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": ""Tahoma"",
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Size"",
              ""Type"": ""System.Decimal"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": 8.25,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Bold"",
              ""Type"": ""System.Boolean"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": false,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Italic"",
              ""Type"": ""System.Boolean"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": false,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Underline"",
              ""Type"": ""System.Boolean"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": false,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Strikethrough"",
              ""Type"": ""System.Boolean"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": false,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Weight"",
              ""Type"": ""System.Int16"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": 400,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            },
            {
              ""Name"": ""Charset"",
              ""Type"": ""System.Int16"",
              ""Kind"": ""scalar"",
              ""ReadOnly"": false,
              ""AllowedValues"": null,
              ""SetterStatus"": ""DescriptorCandidateUnverified"",
              ""Value"": 0,
              ""Display"": null,
              ""Digest"": null,
              ""Error"": null,
              ""NumIndices"": 0,
              ""Members"": null
            }
          ]
        },
        {
          ""Name"": ""Picture"",
          ""Type"": ""System.Drawing.Bitmap"",
          ""Kind"": ""scalar"",
          ""ReadOnly"": false,
          ""AllowedValues"": null,
          ""SetterStatus"": ""DescriptorCandidateUnverified"",
          ""Value"": null,
          ""Display"": null,
          ""Digest"": null,
          ""Error"": null,
          ""NumIndices"": 0,
          ""Members"": null
        }
      ]
    }
  ]
}");

        [TestMethod, TestCategory("Unit")]
        public void PublicationRetainsStrictFailureForCapturedReservedGetterErrorsBeforeReaderCorrection()
        {
            // Historical source data had two unavailable write-only alias getter errors.
            Assert.ThrowsException<InvalidOperationException>(() =>
                VbeProjectComponents.RequirePublicationDesignerReadable(CapturedNativeFontTree()));
        }

        private static List<IDictionary<string, object>> CapturedOwnerProperties(Dictionary<string, object> tree)
        {
            var owners = new List<IDictionary<string, object>> { tree };
            owners.AddRange(((IEnumerable)tree["Controls"]).Cast<IDictionary<string, object>>());
            return owners;
        }

        private static void AssertActualCanonicalFontShapes(Dictionary<string, object> tree)
        {
            var owners = CapturedOwnerProperties(tree);
            Assert.AreEqual(2, owners.Count);
            foreach (var owner in owners)
            {
                var rows = ((IEnumerable)owner["Properties"]).Cast<IDictionary<string, object>>().ToArray();
                var font = rows.Single(x => (string)x["Name"] == "Font");
                Assert.IsNull(font["Error"]);
                var members = ((IEnumerable)font["Members"]).Cast<IDictionary<string, object>>().ToArray();
                Assert.IsTrue(members.All(x => x["Error"] == null));
                Assert.AreEqual("Tahoma", members.Single(x => (string)x["Name"] == "Name")["Value"]);
                Assert.AreEqual(8.25m, Convert.ToDecimal(members.Single(x => (string)x["Name"] == "Size")["Value"]));
                Assert.IsTrue(members.Any(x => (string)x["Name"] == "Bold"));
                Assert.IsTrue(members.Any(x => (string)x["Name"] == "Italic"));
                Assert.IsTrue(members.Any(x => (string)x["Name"] == "Underline"));
            }
            var managed = ((IEnumerable)owners[0]["Properties"]).Cast<IDictionary<string, object>>().Single(x => (string)x["Name"] == "Font");
            var ole = ((IEnumerable)owners[1]["Properties"]).Cast<IDictionary<string, object>>().Single(x => (string)x["Name"] == "Font");
            var managedMembers = ((IEnumerable)managed["Members"]).Cast<IDictionary<string, object>>().ToArray();
            var oleMembers = ((IEnumerable)ole["Members"]).Cast<IDictionary<string, object>>().ToArray();
            Assert.AreEqual(16, managedMembers.Length); Assert.AreEqual(8, oleMembers.Length);
            Assert.IsTrue(managedMembers.Any(x => (string)x["Name"] == "Strikeout"));
            Assert.IsTrue(oleMembers.Any(x => (string)x["Name"] == "Strikethrough"));
        }
        private sealed class CapturedReservedFontDescriptor : System.ComponentModel.PropertyDescriptor
        {
            internal int GetterEntries;
            internal CapturedReservedFontDescriptor() : base("_Font_Reserved", null) { }
            public override Type ComponentType => typeof(object);
            public override Type PropertyType => typeof(System.Drawing.Font);
            public override bool IsReadOnly => false;
            public override bool CanResetValue(object component) => false;
            public override bool ShouldSerializeValue(object component) => false;
            public override object GetValue(object component) { GetterEntries++; throw new InvalidOperationException("Native write-only alias getter must not be invoked."); }
            public override void SetValue(object component, object value) { throw new InvalidOperationException("No native setter in this test."); }
            public override void ResetValue(object component) { throw new InvalidOperationException("No native reset in this test."); }
        }
        private static object CapturedReservedWriteOnlyMetadata() => new
        {
            Property = "_Font_Reserved",
            Discovery = "IProvideClassInfo/ITypeInfo",
            MetadataComplete = true,
            SetterDeclared = true,
            Interfaces = new[] { "Captured MSForms ILabelControl/_UserForm" },
            Accessors = new[] { new { Name = "_Font_Reserved", DispId = 2147483135, InvocationKind = "INVOKE_PROPERTYPUT" } },
            Errors = new string[0]
        };
        private static void ApplyActualReservedReaderClassification(Dictionary<string, object> tree)
        {
            foreach (var owner in CapturedOwnerProperties(tree))
            {
                var properties = ((IEnumerable)owner["Properties"]).Cast<IDictionary<string, object>>().ToArray();
                int index = Array.FindIndex(properties, p => (string)p["Name"] == "_Font_Reserved");
                Assert.IsTrue(index >= 0);
                var descriptor = new CapturedReservedFontDescriptor();
                var info = new VbePropertyInfo { Name = descriptor.Name, Type = descriptor.PropertyType.FullName, ReadOnly = false };
                Assert.IsTrue(VbeForms.TryDescribeReservedFontWriteOnly(descriptor, CapturedReservedWriteOnlyMetadata(), info));
                Assert.AreEqual(0, descriptor.GetterEntries);
                Assert.AreEqual("writeOnly", info.Kind); Assert.AreEqual("GetterUnavailable", info.SetterStatus);
                Assert.IsNull(info.Error); Assert.IsNull(info.Value); Assert.IsNull(info.Members);
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                properties[index] = json.Deserialize<Dictionary<string, object>>(json.Serialize(info));
                owner["Properties"] = properties.Cast<object>().ToArray();
            }
        }
        [TestMethod, TestCategory("Unit")]
        public void PublicationCapturedManagedAndOleFontsPassAfterActualWriteOnlyReaderClassification()
        {
            var tree = CapturedNativeFontTree(); AssertActualCanonicalFontShapes(tree);
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            string fontsBefore = json.Serialize(CapturedOwnerProperties(tree).Select(o =>
                ((IEnumerable)o["Properties"]).Cast<IDictionary<string, object>>().Single(p => (string)p["Name"] == "Font")).ToArray());
            ApplyActualReservedReaderClassification(tree);
            VbeProjectComponents.RequirePublicationDesignerReadable(tree);
            AssertActualCanonicalFontShapes(tree);
            string fontsAfter = json.Serialize(CapturedOwnerProperties(tree).Select(o =>
                ((IEnumerable)o["Properties"]).Cast<IDictionary<string, object>>().Single(p => (string)p["Name"] == "Font")).ToArray());
            Assert.AreEqual(fontsBefore, fontsAfter, "The reserved alias classifier cannot alter actual Font values or members.");
        }
        [DataTestMethod, TestCategory("Unit")]
        [DataRow("font"), DataRow("font-member"), DataRow("picture")]
        public async Task PublicationCapturedPersistentFontOrPictureErrorStillRefusesBeforeExportOrNativeCreation(string fault)
        {
            using (var f = new Fixture())
            {
                var tree = CapturedNativeFontTree(); ApplyActualReservedReaderClassification(tree);
                var properties = ((IEnumerable)tree["Properties"]).Cast<IDictionary<string, object>>().ToList();
                var font = properties.Single(p => (string)p["Name"] == "Font");
                if (fault == "font") font["Error"] = "Canonical font getter unavailable";
                else if (fault == "font-member") ((IEnumerable)font["Members"]).Cast<IDictionary<string, object>>().Single(m => (string)m["Name"] == "Size")["Error"] = "Actual font size unavailable";
                else
                {
                    var picture = properties.SingleOrDefault(p => (string)p["Name"] == "Picture");
                    if (picture == null) { picture = new Dictionary<string, object> { ["Name"] = "Picture" }; properties.Add(picture); tree["Properties"] = properties.Cast<object>().ToArray(); }
                    picture["Error"] = "Actual picture persistence digest unavailable";
                }
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                f.Service.PublicationFormTree = (project, form) => tree;
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Service.PublishSolidWorksMacroAsync(f.Request()));
                Assert.AreEqual(0, f.CreateAttempts); Assert.AreEqual(0, f.SaveAttempts); Assert.AreEqual(0, f.GeneralReads);
                Assert.IsTrue(f.Source.VBComponents.Items.All(c => c.ExportAttempts == 0));
                Assert.IsFalse(File.Exists(f.Path)); Assert.AreEqual(1, f.Vbe.VBProjects.Count);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationCapturedManagedAndOleFontShapesReachFullCopyWithActualAliasClassification()
        {
            using (var f = new Fixture())
            {
                var tree = CapturedNativeFontTree(); ApplyActualReservedReaderClassification(tree);
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                f.Service.PublicationFormTree = (project, form) => tree;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsTrue(result.Verified, result.Error); Assert.IsTrue(result.OriginalPreserved); Assert.IsTrue(result.Terminal); Assert.IsFalse(result.Uncertain);
                Assert.AreEqual(1, f.CreateAttempts); Assert.AreEqual(1, f.SaveAttempts); Assert.AreEqual(5, f.GeneralReads);
                Assert.IsTrue(File.Exists(System.IO.Path.Combine(result.StagingPath, "Form1.frx")), result.Error);
                Assert.IsFalse(result.RetryAllowed);
            }
        }
        [DataTestMethod, TestCategory("Unit"), DataRow("name"), DataRow("type"), DataRow("code"), DataRow("designer")]
        public void PublicationImportDiagnosticReportsExactMismatchAndOwnedImmutableObservations(string field)
        {
            using (var f = new Fixture())
            {
                var left = new VbeProjectComponents.PublicationComponent
                {
                    Name = "Form1",
                    Type = 3,
                    Code = "Option Explicit\r\n",
                    FormVersion = "tree-source",
                    DesignerJson = "{\"Properties\":[]}"
                };
                var right = new VbeProjectComponents.PublicationComponent
                {
                    Name = left.Name,
                    Type = left.Type,
                    Code = left.Code,
                    FormVersion = left.FormVersion,
                    DesignerJson = left.DesignerJson
                };
                if (field == "name") right.Name = "Form2";
                if (field == "type") right.Type = 1;
                if (field == "code") right.Code = "\r\n" + right.Code;
                if (field == "designer") right.FormVersion = "tree-destination";
                var result = VbeProjectComponents.CapturePublicationImportMismatch(left, right, f.Root, 1);
                Assert.AreEqual(field != "name", result.NameEqual); Assert.AreEqual(field != "type", result.TypeEqual);
                Assert.AreEqual(field != "code", result.CodeEqual); Assert.AreEqual(field != "designer", result.FormVersionEqual);
                Assert.IsNull(result.DiagnosticError); Assert.AreEqual(0, result.ExpectedLeadingEmptyLines);
                Assert.AreEqual(field == "code" ? 1 : 0, result.ActualLeadingEmptyLines);
                Assert.AreEqual(field == "code" ? 1 : -1, result.FirstDifferentCodeLine);
                Assert.IsFalse(System.IO.Path.IsPathRooted(result.ExpectedRelativeFile));
                string observation = File.ReadAllText(System.IO.Path.Combine(f.Root, result.ExpectedRelativeFile));
                StringAssert.Contains(observation, "Option Explicit"); StringAssert.Contains(observation, "Properties");
                Assert.AreEqual(64, result.ExpectedFileSha256.Length); Assert.AreEqual(64, result.ActualCodeSha256.Length);
                right.Code = "changed later"; Assert.AreEqual(field != "code", result.CodeEqual);
                Assert.AreEqual(0, f.CreateAttempts); Assert.AreEqual(0, f.SaveAttempts); Assert.AreEqual(0, f.GeneralReads);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public void PublicationImportDiagnosticNeverOverwritesAndDoesNotHidePrimaryMismatch()
        {
            using (var f = new Fixture())
            {
                var left = new VbeProjectComponents.PublicationComponent { Name = "Form1", Type = 3, Code = "source", FormVersion = "one" };
                var right = new VbeProjectComponents.PublicationComponent { Name = "Form1", Type = 3, Code = "destination", FormVersion = "two" };
                var first = VbeProjectComponents.CapturePublicationImportMismatch(left, right, f.Root, 1);
                byte[] original = File.ReadAllBytes(System.IO.Path.Combine(f.Root, first.ExpectedRelativeFile));
                left.Code = "second source";
                var second = VbeProjectComponents.CapturePublicationImportMismatch(left, right, f.Root, 1);
                Assert.IsNotNull(second.DiagnosticError); Assert.IsFalse(second.CodeEqual); Assert.IsFalse(second.FormVersionEqual);
                CollectionAssert.AreEqual(original, File.ReadAllBytes(System.IO.Path.Combine(f.Root, first.ExpectedRelativeFile)));
                Assert.AreEqual(0, f.CreateAttempts); Assert.AreEqual(0, f.SaveAttempts);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public void PublicationImportDiagnosticRejectsOversizedLocalObservationWithoutWriting()
        {
            using (var f = new Fixture())
            {
                var left = new VbeProjectComponents.PublicationComponent { Name = "Form1", Type = 3, Code = new string('x', 16 * 1024 * 1024 + 1) };
                var right = new VbeProjectComponents.PublicationComponent { Name = "Form1", Type = 3, Code = "destination" };
                var result = VbeProjectComponents.CapturePublicationImportMismatch(left, right, f.Root, 1);
                Assert.IsNotNull(result.DiagnosticError); Assert.IsFalse(result.CodeEqual);
                Assert.AreEqual(0, Directory.GetFiles(f.Root, "import-mismatch-*.json").Length);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationImportDiagnosticPreservesUncertainTerminalFailureAndZeroSave()
        {
            using (var f = new Fixture())
            {
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                f.Service.PublicationFormTree = (project, form) => new Dictionary<string, object>
                {
                    ["TreeVersion"] = project == "Published" ? "destination-tree" : "source-tree",
                    ["Properties"] = new[] { new { Name = "Caption", Value = project, Error = (string)null } },
                    ["Controls"] = new object[0]
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Terminal); Assert.IsTrue(result.Uncertain);
                Assert.IsTrue(result.MutationInvoked); Assert.IsTrue(result.DestinationCreated); Assert.IsTrue(result.OriginalPreserved);
                Assert.AreEqual(0, f.SaveAttempts); Assert.IsNull(result.Save); Assert.IsFalse(result.RetryAllowed);
                Assert.IsNotNull(result.ImportMismatch); Assert.IsFalse(result.ImportMismatch.FormVersionEqual);
                Assert.IsNull(result.ImportMismatch.DiagnosticError);
                Assert.IsTrue(File.Exists(System.IO.Path.Combine(result.StagingPath, result.ImportMismatch.ActualRelativeFile)));
                StringAssert.Contains(result.Error, "Imported component code or complete designer/picture state differs.");
            }
        }
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationImportDiagnosticDoesNotBypassUnreadableSourceDesigner()
        {
            using (var f = new Fixture())
            {
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                f.Service.PublicationFormTree = (project, form) => new Dictionary<string, object>
                {
                    ["TreeVersion"] = "unreadable",
                    ["Properties"] = new[] { new { Name = "Picture", Error = "Missing persisted picture" } }
                };
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Service.PublishSolidWorksMacroAsync(f.Request()));
                Assert.AreEqual(0, f.CreateAttempts); Assert.AreEqual(0, f.SaveAttempts); Assert.AreEqual(0, f.GeneralReads);
                Assert.AreEqual(0, Directory.GetFiles(f.Root, "import-mismatch-*.json", SearchOption.AllDirectories).Length);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationMatchedImportsCreateNoMismatchObservationFiles()
        {
            using (var f = new Fixture())
            {
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsTrue(result.Verified, result.Error); Assert.IsNull(result.ImportMismatch);
                Assert.AreEqual(0, Directory.GetFiles(result.StagingPath, "import-mismatch-*.json").Length);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public async Task PublicationLocalMismatchWriteFailurePreservesPrimaryFailureAndZeroSave()
        {
            using (var f = new Fixture())
            {
                f.Source.VBComponents.Items.Add(new Component("Form1", 3));
                string retained = null;
                f.Service.PublicationFormTree = (project, form) =>
                {
                    if (project == "Published")
                    {
                        string staging = Directory.GetDirectories(f.Root, ".vbai-publication-*").Single();
                        retained = System.IO.Path.Combine(staging, "import-mismatch-0002-expected.json");
                        File.WriteAllText(retained, "Existing observation must remain unchanged");
                    }
                    return new Dictionary<string, object> { ["TreeVersion"] = project == "Published" ? "actual" : "expected", ["Properties"] = new[] { new { Name = "Caption", Value = project, Error = (string)null } }, ["Controls"] = new object[0] };
                };
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.IsFalse(result.Verified); Assert.IsTrue(result.Terminal); Assert.IsTrue(result.Uncertain);
                Assert.AreEqual(0, f.SaveAttempts); Assert.IsNull(result.Save); Assert.IsFalse(result.RetryAllowed);
                StringAssert.Contains(result.Error, "Imported component code or complete designer/picture state differs.");
                Assert.IsNotNull(result.ImportMismatch); Assert.IsNotNull(result.ImportMismatch.DiagnosticError);
                Assert.IsNull(result.ImportMismatch.ExpectedFileSha256);
                Assert.AreEqual("Existing observation must remain unchanged", File.ReadAllText(retained));
            }
        }
        private static Dictionary<string, object> Native09PublicationPair(bool destination)
        {
            string raw = destination ? @"{""Name"":""Q020PublishedForm"",""Type"":3,""Code"":""\r\nOption Explicit\r\nPrivate Sub OwnedMarker()\r\n    Dim marker As Long\r\n    marker = 20\r\nEnd Sub"",""FormVersion"":""16970f550b00c934ef9acd96df238beeafba5601ebfa3b1b6ecb15d257f4ebae"",""Designer"":{""Project"":""diagnostic_published_native"",""Form"":""Q020PublishedForm"",""FormVersion"":""16970f550b00c934ef9acd96df238beeafba5601ebfa3b1b6ecb15d257f4ebae"",""TreeVersion"":""16970f550b00c934ef9acd96df238beeafba5601ebfa3b1b6ecb15d257f4ebae"",""NodeCount"":1,""Properties"":[{""Name"":""ActiveControl"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Application"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VBE"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""BackColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483633,""Display"":""Control"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483630,""Display"":""ControlText"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderStyle"",""Type"":""1767013632_fmBorderStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Single""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanPaste"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanRedo"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanUndo"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Controls"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_NewEnum"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""Cycle"",""Type"":""1767013720_fmCycle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""AllForms"",""CurrentForm""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Font_Reserved"",""Type"":""System.Drawing.Font"",""Kind"":""writeOnly"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""GetterUnavailable"",""Value"":null,""Display"":""(write-only COM alias)"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Font"",""Type"":""System.Drawing.Font"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""Font"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""FontFamily"",""Type"":""System.Drawing.FontFamily"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":""[FontFamily: Name=Tahoma]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Bold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GdiCharSet"",""Type"":""System.Byte"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GdiVerticalFont"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Italic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OriginalFontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Strikeout"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Underline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Style"",""Type"":""System.Drawing.FontStyle"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""Regular"",""Bold"",""Italic"",""Underline"",""Strikeout""],""SetterStatus"":""DescriptorReadOnly"",""Value"":""Regular"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Size"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SizeInPoints"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Unit"",""Type"":""System.Drawing.GraphicsUnit"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""World"",""Display"",""Pixel"",""Point"",""Inch"",""Document"",""Millimeter""],""SetterStatus"":""DescriptorReadOnly"",""Value"":""Point"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":14,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""IsSystemFont"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SystemFontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":"""",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""ForeColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483630,""Display"":""ControlText"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InsideHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":150.75,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InsideWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":228,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""KeepScrollBarsVisible"",""Type"":""1767013896_fmScrollBars"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Horizontal"",""Vertical"",""Both""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":3,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MouseIcon"",""Type"":""System.Drawing.Icon"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Application"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":5,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VBE"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""MousePointer"",""Type"":""1767013984_fmMousePointer"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Default"",""Arrow"",""Cross"",""IBeam"",""SizeNESW"",""SizeNS"",""SizeNWSE"",""SizeWE"",""UpArrow"",""HourGlass"",""NoDrop"",""AppStarting"",""Help"",""SizeAll"",""Custom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureAlignment"",""Type"":""1767014072_fmPictureAlignment"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""TopLeft"",""TopRight"",""Center"",""BottomLeft"",""BottomRight""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Picture"",""Type"":""stdole.IPictureDisp"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",""Digest"":""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureSizeMode"",""Type"":""1767014160_fmPictureSizeMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Clip"",""Stretch"",""Zoom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureTiling"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollBars"",""Type"":""1767013896_fmScrollBars"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Horizontal"",""Vertical"",""Both""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollLeft"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollTop"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Selected"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_NewEnum"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""SpecialEffect"",""Type"":""1767014248_fmSpecialEffect"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Flat"",""Raised"",""Sunken"",""Etched"",""Bump""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VerticalScrollBarSide"",""Type"":""1767014424_fmVerticalScrollBarSide"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Right"",""Left""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Zoom"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":100,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""DesignMode"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ShowToolbox"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ShowGridDots"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SnapToGrid"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GridX"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GridY"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""DrawBuffer"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":32000,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":""Q020PublishedForm"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Caption"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""UserForm1"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Left"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Top"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Width"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":240,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":180,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Enabled"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Tag"",""Type"":null,""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""HelpContextID"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""WhatsThisButton"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""WhatsThisHelp"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RightToLeft"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""StartUpPosition"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ShowModal"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}],""Controls"":[{""Path"":""Controls/Q020Label"",""Name"":""Q020Label"",""Kind"":""Control"",""Type"":""Label"",""Properties"":[{""Name"":""Cancel"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""BlockedNativeSetterFailure"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ControlSource"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ControlTipText"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Default"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":24,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""HelpContextID"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InSelection"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""LayoutEffect"",""Type"":""1841362504_fmLayoutEffect"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""None"",""Initiate"",""Respond""],""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Left"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":12,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020Label"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldLeft"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldTop"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Object"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RowSource"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RowSourceType"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TabIndex"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TabStop"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Tag"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Top"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":12,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BoundValue"",""Type"":""System.Windows.Forms.ComponentModel.Com2Interop.Com2Variant"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Visible"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Width"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":160,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""AutoSize"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BackColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [Control]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BackStyle"",""Type"":""1841362680_fmBackStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Transparent"",""Opaque""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Opaque"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [WindowFrame]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderStyle"",""Type"":""1841362768_fmBorderStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Single""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""None"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Caption"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Enabled"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Font_Reserved"",""Type"":""System.Drawing.Font"",""Kind"":""writeOnly"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""GetterUnavailable"",""Value"":null,""Display"":""(write-only COM alias)"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Font"",""Type"":""System.Drawing.Font"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Size"",""Type"":""System.Decimal"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Bold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Italic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Underline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Strikethrough"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Weight"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":400,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Charset"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""FontItalic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontBold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontSize"",""Type"":""System.Decimal"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontStrikethru"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontUnderline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ForeColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [ControlText]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MouseIcon"",""Type"":""System.Drawing.Icon"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MousePointer"",""Type"":""1841362944_fmMousePointer"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Default"",""Arrow"",""Cross"",""IBeam"",""SizeNESW"",""SizeNS"",""SizeNWSE"",""SizeWE"",""UpArrow"",""HourGlass"",""NoDrop"",""AppStarting"",""Help"",""SizeAll"",""Custom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Default"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Picture"",""Type"":""System.Drawing.Bitmap"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PicturePosition"",""Type"":""1841363032_fmPicturePosition"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""LeftTop"",""LeftCenter"",""LeftBottom"",""RightTop"",""RightCenter"",""RightBottom"",""AboveLeft"",""AboveCenter"",""AboveRight"",""BelowLeft"",""BelowCenter"",""BelowRight"",""Center""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""AboveCenter"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SpecialEffect"",""Type"":""1841363120_fmSpecialEffect"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Flat"",""Raised"",""Sunken"",""Etched"",""Bump""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Flat"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TextAlign"",""Type"":""1841363208_fmTextAlign"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Left"",""Center"",""Right""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Left"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""WordWrap"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Accelerator"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontWeight"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":400,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Value"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}],""Children"":[]}]}}" : @"{""Name"":""Q020PublishedForm"",""Type"":3,""Code"":""Option Explicit\r\nPrivate Sub OwnedMarker()\r\n    Dim marker As Long\r\n    marker = 20\r\nEnd Sub"",""FormVersion"":""2ffc87e6f4ab6f8e822ea9cdde2db93ffd9228389a6194896ea113d8a191d3e7"",""Designer"":{""Project"":""Project1"",""Form"":""Q020PublishedForm"",""FormVersion"":""2ffc87e6f4ab6f8e822ea9cdde2db93ffd9228389a6194896ea113d8a191d3e7"",""TreeVersion"":""2ffc87e6f4ab6f8e822ea9cdde2db93ffd9228389a6194896ea113d8a191d3e7"",""NodeCount"":1,""Properties"":[{""Name"":""ActiveControl"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Application"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VBE"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""BackColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483633,""Display"":""Control"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483630,""Display"":""ControlText"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderStyle"",""Type"":""1767013632_fmBorderStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Single""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanPaste"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanRedo"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""CanUndo"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Controls"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_NewEnum"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""Cycle"",""Type"":""1767013720_fmCycle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""AllForms"",""CurrentForm""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Font_Reserved"",""Type"":""System.Drawing.Font"",""Kind"":""writeOnly"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""GetterUnavailable"",""Value"":null,""Display"":""(write-only COM alias)"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Font"",""Type"":""System.Drawing.Font"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""Font"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""FontFamily"",""Type"":""System.Drawing.FontFamily"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":""[FontFamily: Name=Tahoma]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Bold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GdiCharSet"",""Type"":""System.Byte"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GdiVerticalFont"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Italic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OriginalFontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Strikeout"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Underline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Style"",""Type"":""System.Drawing.FontStyle"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""Regular"",""Bold"",""Italic"",""Underline"",""Strikeout""],""SetterStatus"":""DescriptorReadOnly"",""Value"":""Regular"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Size"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SizeInPoints"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Unit"",""Type"":""System.Drawing.GraphicsUnit"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""World"",""Display"",""Pixel"",""Point"",""Inch"",""Document"",""Millimeter""],""SetterStatus"":""DescriptorReadOnly"",""Value"":""Point"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":14,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""IsSystemFont"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SystemFontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":"""",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""ForeColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2147483630,""Display"":""ControlText"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InsideHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":150.75,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InsideWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":228,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""KeepScrollBarsVisible"",""Type"":""1767013896_fmScrollBars"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Horizontal"",""Vertical"",""Both""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":3,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MouseIcon"",""Type"":""System.Drawing.Icon"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Application"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":5,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VBE"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""MousePointer"",""Type"":""1767013984_fmMousePointer"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Default"",""Arrow"",""Cross"",""IBeam"",""SizeNESW"",""SizeNS"",""SizeNWSE"",""SizeWE"",""UpArrow"",""HourGlass"",""NoDrop"",""AppStarting"",""Help"",""SizeAll"",""Custom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureAlignment"",""Type"":""1767014072_fmPictureAlignment"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""TopLeft"",""TopRight"",""Center"",""BottomLeft"",""BottomRight""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Picture"",""Type"":""stdole.IPictureDisp"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",""Digest"":""ole-persist-v1:1:212:212:4fc22834f3490407ec073f2139da3260453c86b545f11efee9a1aa9a9b7330e1"",""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureSizeMode"",""Type"":""1767014160_fmPictureSizeMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Clip"",""Stretch"",""Zoom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PictureTiling"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollBars"",""Type"":""1767013896_fmScrollBars"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Horizontal"",""Vertical"",""Both""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollLeft"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollTop"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ScrollWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Selected"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":""__ComObject"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Count"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_NewEnum"",""Type"":""System.Object"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""SpecialEffect"",""Type"":""1767014248_fmSpecialEffect"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Flat"",""Raised"",""Sunken"",""Etched"",""Bump""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""VerticalScrollBarSide"",""Type"":""1767014424_fmVerticalScrollBarSide"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Right"",""Left""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Zoom"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":100,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""DesignMode"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ShowToolbox"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ShowGridDots"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SnapToGrid"",""Type"":""1767014512_fmMode"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Off"",""Inherit"",""On""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":-2,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GridX"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""GridY"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""DrawBuffer"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":32000,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":""Q020PublishedForm"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Left"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Top"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Width"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":240,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":180,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Caption"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""UserForm1"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Enabled"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Visible"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Tag"",""Type"":null,""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""StartUpPosition"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":1,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RightToLeft"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":null,""AllowedValues"":null,""SetterStatus"":""Unknown"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}],""Controls"":[{""Path"":""Controls/Q020Label"",""Name"":""Q020Label"",""Kind"":""Control"",""Type"":""Label"",""Properties"":[{""Name"":""Cancel"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""BlockedNativeSetterFailure"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ControlSource"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ControlTipText"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Default"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Height"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":24,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""HelpContextID"",""Type"":""System.Int32"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""InSelection"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""LayoutEffect"",""Type"":""1841362504_fmLayoutEffect"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":[""None"",""Initiate"",""Respond""],""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Left"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":12,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020Label"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldHeight"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldLeft"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldTop"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""OldWidth"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Object"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Parent"",""Type"":""System.Windows.Forms.UnsafeNativeMethods+IDispatch"",""Kind"":""object"",""ReadOnly"":true,""AllowedValues"":null,""SetterStatus"":""DescriptorReadOnly"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RowSource"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""RowSourceType"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TabIndex"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TabStop"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Tag"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Top"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":12,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BoundValue"",""Type"":""System.Windows.Forms.ComponentModel.Com2Interop.Com2Variant"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Visible"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Width"",""Type"":""System.Single"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":160,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""AutoSize"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BackColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [Control]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BackStyle"",""Type"":""1841362680_fmBackStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Transparent"",""Opaque""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Opaque"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [WindowFrame]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""BorderStyle"",""Type"":""1841362768_fmBorderStyle"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""None"",""Single""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""None"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Caption"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Enabled"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Font_Reserved"",""Type"":""System.Drawing.Font"",""Kind"":""writeOnly"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""GetterUnavailable"",""Value"":null,""Display"":""(write-only COM alias)"",""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Font"",""Type"":""System.Drawing.Font"",""Kind"":""object"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":[{""Name"":""Name"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Size"",""Type"":""System.Decimal"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Bold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Italic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Underline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Strikethrough"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Weight"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":400,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Charset"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":0,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}]},{""Name"":""FontItalic"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontBold"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontName"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Tahoma"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontSize"",""Type"":""System.Decimal"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":8.25,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontStrikethru"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontUnderline"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":false,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""ForeColor"",""Type"":""System.Drawing.Color"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Color [ControlText]"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MouseIcon"",""Type"":""System.Drawing.Icon"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""MousePointer"",""Type"":""1841362944_fmMousePointer"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Default"",""Arrow"",""Cross"",""IBeam"",""SizeNESW"",""SizeNS"",""SizeNWSE"",""SizeWE"",""UpArrow"",""HourGlass"",""NoDrop"",""AppStarting"",""Help"",""SizeAll"",""Custom""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Default"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Picture"",""Type"":""System.Drawing.Bitmap"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""PicturePosition"",""Type"":""1841363032_fmPicturePosition"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""LeftTop"",""LeftCenter"",""LeftBottom"",""RightTop"",""RightCenter"",""RightBottom"",""AboveLeft"",""AboveCenter"",""AboveRight"",""BelowLeft"",""BelowCenter"",""BelowRight"",""Center""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""AboveCenter"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""SpecialEffect"",""Type"":""1841363120_fmSpecialEffect"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Flat"",""Raised"",""Sunken"",""Etched"",""Bump""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Flat"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""TextAlign"",""Type"":""1841363208_fmTextAlign"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":[""Left"",""Center"",""Right""],""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Left"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""WordWrap"",""Type"":""System.Boolean"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":true,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""Accelerator"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":null,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""FontWeight"",""Type"":""System.Int16"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":400,""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null},{""Name"":""_Value"",""Type"":""System.String"",""Kind"":""scalar"",""ReadOnly"":false,""AllowedValues"":null,""SetterStatus"":""DescriptorCandidateUnverified"",""Value"":""Q020 native owned"",""Display"":null,""Digest"":null,""Error"":null,""NumIndices"":0,""Members"":null}],""Children"":[]}]}}";
            return new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Dictionary<string, object>>(raw);
        }
        private static readonly string Native09ExportHeader = @"VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Q020PublishedForm 
   Caption         =   ""UserForm1""
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   ""Q020PublishedForm.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
        [DataTestMethod, TestCategory("Unit"), DataRow("none"), DataRow("HelpContextID"), DataRow("ShowModal"),
         DataRow("WhatsThisButton"), DataRow("WhatsThisHelp"), DataRow("Font"), DataRow("Picture"), DataRow("unknown"), DataRow("Visible"), DataRow("control-visible"), DataRow("code-body")]
        public async Task PublicationNative09PairPreservesContentAndRefusesNondefaultOrRealResourceChanges(string change)
        {
            using (var f = new Fixture())
            {
                var source = Native09PublicationPair(false); var imported = Native09PublicationPair(true);
                string formName = (string)source["Name"];
                var sourceTree = (Dictionary<string, object>)source["Designer"];
                var actualTree = (Dictionary<string, object>)imported["Designer"];
                var properties = ((IEnumerable)actualTree["Properties"]).Cast<Dictionary<string, object>>().ToList();
                if (change == "HelpContextID") properties.Single(x => (string)x["Name"] == change)["Value"] = 42;
                else if (new[] { "ShowModal", "WhatsThisButton", "WhatsThisHelp" }.Contains(change))
                    properties.Single(x => (string)x["Name"] == change)["Value"] = !(bool)properties.Single(x => (string)x["Name"] == change)["Value"];
                else if (change == "Font")
                    ((IEnumerable)properties.Single(x => (string)x["Name"] == "Font")["Members"]).Cast<Dictionary<string, object>>().Single(x => (string)x["Name"] == "Name")["Value"] = "Arial";
                else if (change == "Picture") properties.Single(x => (string)x["Name"] == "Picture")["Digest"] = "different-real-picture";
                else if (change == "unknown") properties.Add(new Dictionary<string, object> { ["Name"] = "UnknownPersistentProperty", ["Value"] = 42, ["Error"] = null });
                if (change == "Visible")
                    ((IEnumerable)sourceTree["Properties"]).Cast<Dictionary<string, object>>().Single(x => (string)x["Name"] == "Visible")["Value"] = false;
                if (change == "control-visible")
                {
                    var node = ((IEnumerable)actualTree["Controls"]).Cast<Dictionary<string, object>>().Single();
                    ((IEnumerable)node["Properties"]).Cast<Dictionary<string, object>>().Single(x => (string)x["Name"] == "Visible")["Value"] = false;
                }
                if (change == "code-body") imported["Code"] = ((string)imported["Code"]).Replace("marker = 20", "marker = 21");
                actualTree["Properties"] = properties.Cast<object>().ToArray();
                var form = new Component(formName, 3) { FormExportHeader = Native09ExportHeader };
                form.CodeModule.Source = (string)source["Code"];
                f.Source.VBComponents.Items.Add(form);
                f.AfterCreate = () =>
                {
                    f.Target.VBComponents.ImportedCodeOverrides[formName] = (string)imported["Code"];
                    f.Target.VBComponents.FormExportHeaderOverrides[formName] = Native09ExportHeader;
                };
                f.Service.PublicationFormTree = (project, name) => project == "Published" ? actualTree : sourceTree;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request());
                Assert.AreEqual(change == "none", result.Verified, result.Error);
                Assert.AreEqual(change == "none" ? 1 : 0, f.SaveAttempts);
                Assert.AreEqual((string)source["Code"], form.CodeModule.Source, "Original native source code must remain exact.");
                Assert.IsTrue(result.OriginalPreserved); Assert.IsFalse(result.RetryAllowed);
                if (change == "none") Assert.AreEqual(form.CodeModule.Source, f.Target.VBComponents.Items.Single(x => x.Name == formName).CodeModule.Source);
                else { Assert.IsTrue(result.Uncertain); Assert.IsNull(result.Save); }
            }
        }
        [DataTestMethod, TestCategory("Unit"), DataRow("source"), DataRow("code"), DataRow("tree"), DataRow("canonical"), DataRow("mode"), DataRow("revoked"), DataRow("delete-throws"), DataRow("runtime-undo")]
        public async Task PublicationPrefixMutationRevalidatesAfterClaimAndNeverReplays(string fault)
        {
            using (var f = new Fixture())
            {
                var left = Native09PublicationPair(false); var right = Native09PublicationPair(true);
                var sourceTree = (Dictionary<string, object>)left["Designer"]; var importedTree = (Dictionary<string, object>)right["Designer"];
                string name = (string)left["Name"];
                var original = new Component(name, 3) { FormExportHeader = Native09ExportHeader };
                original.CodeModule.Source = (string)left["Code"]; f.Source.VBComponents.Items.Add(original);
                f.AfterCreate = () => { f.Target.VBComponents.ImportedCodeOverrides[name] = (string)right["Code"]; f.Target.VBComponents.FormExportHeaderOverrides[name] = Native09ExportHeader; };
                f.Service.PublicationFormTree = (project, form) => project == "Published" ? importedTree : sourceTree;
                bool revoked = false; PublicationCodeModule liveCode = null;
                var result = (VbeProjectComponents.SolidWorksMacroPublicationResult)await f.Service.PublishSolidWorksMacroAsync(f.Request(),
                    shared => { if (revoked) throw new UnauthorizedAccessException("Revoked publication grant"); },
                    claim =>
                    {
                        if (claim.Phase != "BeforeImportedCodePrefixReconciliation") return;
                        var live = f.Target.VBComponents.Items.Single(x => x.Name == name); liveCode = live.CodeModule;
                        if (fault == "source") original.CodeModule.Source = "Changed original";
                        if (fault == "code") live.CodeModule.Source = "Real content inserted before prefix deletion";
                        if (fault == "tree") { importedTree["TreeVersion"] = "Changed local revision"; }
                        if (fault == "canonical") { f.Target.VBComponents.Items.Remove(live); var replacement = new Component(name, 3) { FormExportHeader = Native09ExportHeader }; replacement.CodeModule.Source = live.CodeModule.Source; f.Target.VBComponents.Items.Add(replacement); }
                        if (fault == "mode") f.Target.Mode = 1;
                        if (fault == "revoked") revoked = true;
                        if (fault == "delete-throws") live.CodeModule.ThrowAfterDelete = true;
                        if (fault == "runtime-undo") live.CodeModule.BeforeDelete = () => { importedTree["TreeVersion"] = "New raw baseline after our operation"; ((IEnumerable)importedTree["Properties"]).Cast<Dictionary<string, object>>().Single(p => (string)p["Name"] == "CanUndo")["Value"] = true; };
                    });
                Assert.IsNotNull(liveCode, result.Error);
                Assert.AreEqual(fault == "delete-throws" || fault == "runtime-undo" ? 1 : 0, liveCode.DeleteAttempts, result.Error);
                Assert.AreEqual(fault == "runtime-undo", result.Verified, result.Error);
                Assert.AreEqual(fault == "runtime-undo" ? 1 : 0, f.SaveAttempts);
                Assert.IsFalse(result.RetryAllowed); Assert.IsTrue(result.Terminal);
                if (fault != "runtime-undo") { Assert.IsTrue(result.Uncertain); Assert.IsNull(result.Save); }
                if (fault != "source") Assert.AreEqual((string)left["Code"], original.CodeModule.Source);
            }
        }
        private sealed class Fixture : IDisposable
        {
            internal readonly string Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-publication-" + Guid.NewGuid().ToString("N"));
            internal string Path, ImportFault;
            internal bool CreationUnsettled;
            internal int CreateAttempts, SaveAttempts;
            internal int GeneralReads;
            internal string ConditionalCompilation = "", GeneralOptionsVersion = "native-options-baseline";
            internal Action<int> BeforeGeneralRead, AfterGeneralRead;
            internal Action AfterCreate;
            internal Reference NativeReference;
            internal readonly Editor Vbe = new Editor();
            internal readonly Project Source = new Project { Name = "Source" };
            internal Project Target;
            internal readonly VbeProjectComponents Service;
            internal Fixture()
            {
                Directory.CreateDirectory(Root); Path = System.IO.Path.Combine(Root, "Published.swp");
                Source.VBComponents.Items.Add(new Component("Module1", 1)); Vbe.VBProjects.Add(Source);
                Service = new VbeProjectComponents(Vbe, new VbeForms(Vbe));
                Service.GeneralProjectIdentity = ReferenceEquals;
                Service.SolidWorksSaveProbe = () => new Owner();
                Service.PublicationReadGeneral = async request =>
                {
                    Assert.AreEqual("read_project_general", request.Command);
                    Assert.AreEqual(Source.Name, request.Project);
                    Assert.AreEqual(2, request.ExpectedMode);
                    Assert.AreEqual((string)((dynamic)Service.ProjectProperties(Source.Name)).Version, request.ExpectedProjectVersion);
                    int ordinal = ++GeneralReads;
                    BeforeGeneralRead?.Invoke(ordinal);
                    object snapshot = GeneralSnapshot();
                    await Task.Yield();
                    AfterGeneralRead?.Invoke(ordinal);
                    return snapshot;
                };
                Service.PublicationCreate = (r, a, c, n) =>
                {
                    CreateAttempts++; Target = new Project { Name = "Published", Type = 100, FileName = r.Path };
                    Target.VBComponents.Items.Add(new Component("ThisLibrary", 100));
                    Target.VBComponents.Items.Add(new Component("Published1", 1));
                    if (NativeReference != null) Target.References.Items.Add(NativeReference);
                    Target.VBComponents.Fault = ImportFault; Vbe.VBProjects.Add(Target); File.WriteAllText(r.Path, "Native100");
                    AfterCreate?.Invoke();
                    return Task.FromResult(new VbeProjectComponents.SolidWorksMacroCreationResult
                    {
                        Verified = !CreationUnsettled,
                        Terminal = !CreationUnsettled,
                        CommandEntered = true,
                        OriginalCommandReturned = !CreationUnsettled,
                        DestinationCreated = true,
                        Project = Target.Name,
                        HostPath = r.Path,
                        Uncertain = CreationUnsettled
                    });
                };
                Service.PublicationSave = r => { SaveAttempts++; return Task.FromResult<object>(new { Verified = true, Uncertain = false }); };
            }
            internal Request Request() => new Request
            {
                Project = Source.Name,
                ExpectedMode = 2,
                Path = Path,
                ExpectedProjectVersion = (string)((dynamic)Service.ProjectProperties(Source.Name)).Version
            };
            internal Dictionary<string, object> GeneralSnapshot() => new Dictionary<string, object>
            {
                ["Available"] = true,
                ["Terminal"] = true,
                ["OriginalExecuteReturned"] = true,
                ["DialogClosed"] = true,
                ["Uncertain"] = false,
                ["MutationInvoked"] = false,
                ["CommittedRequested"] = false,
                ["Error"] = null,
                ["OpenAttempts"] = 1,
                ["CancelAttempts"] = 1,
                ["FieldAttempts"] = 0,
                ["OkAttempts"] = 0,
                ["OptionsVersion"] = GeneralOptionsVersion,
                ["Name"] = Source.Name,
                ["Description"] = Source.Description ?? "",
                ["HelpFile"] = Source.HelpFile ?? "",
                ["HelpContextText"] = Source.HelpContextID.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["ConditionalCompilation"] = ConditionalCompilation
            };
            public void Dispose() { Directory.Delete(Root, true); }
        }
        public sealed class Editor { public List<Project> VBProjects { get; } = new List<Project>(); }
        public sealed class Project
        {
            public string Name { get; set; }
            public string Description { get; set; } = "Original description";
            public string FileName { get; set; } = "";
            public string HelpFile { get; set; } = "";
            public int HelpContextID { get; set; }
            public int Type { get; set; } = 101;
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
            public bool Saved { get; set; } = true;
            public Components VBComponents { get; } = new Components();
            public References References { get; } = new References();
        }
        public sealed class Component : VbeProjectComponentsTests.FakeComponent
        {
            public Component(string name, int type) : base(name, type) { }
            public bool Saved { get; set; }
            public bool HasOpenDesigner { get; set; }
            public int UnknownScalar { get; set; } = 1;
            // Instrumentation is a field: VBIDE property snapshots must not observe a fictitious changing property.
            public new int ExportAttempts;
            public string FormExportHeader;
            public new PublicationCodeModule CodeModule { get; } = new PublicationCodeModule();
            public new void Export(string path)
            {
                ExportAttempts++;
                string header = "";
                if (Type == 3)
                {
                    header = FormExportHeader ?? "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + Name + "\r\n   OleObjectBlob = \"" + System.IO.Path.GetFileName(System.IO.Path.ChangeExtension(path, ".frx")) + "\":0000\r\nEnd\r\n";
                    header = System.Text.RegularExpressions.Regex.Replace(header, "(?m)(OleObjectBlob\\s*=\\s*)\"[^\"]+\"", m => m.Groups[1].Value + "\"" + System.IO.Path.GetFileName(System.IO.Path.ChangeExtension(path, ".frx")) + "\"");
                }
                File.WriteAllText(path, header + "Attribute VB_Name = \"" + Name + "\"\r\n" +
                    (Type == 2 ? "Attribute VB_PredeclaredId = True\r\n" : "") + CodeModule.Source);
                if (Type == 3) File.WriteAllBytes(System.IO.Path.ChangeExtension(path, ".frx"), new byte[] { 1, 2, 3, 4 });
            }
        }
        public sealed class PublicationCodeModule
        {
            public string Source { get; set; } = "Option Explicit";
            public bool FailRead;
            public int DeleteAttempts;
            public bool ThrowAfterDelete;
            public Action BeforeDelete;
            public int CountOfLines { get { if (FailRead) throw new InvalidOperationException("Code unavailable"); return Source.Length == 0 ? 0 : Source.Split('\n').Length; } }
            public string Lines(int start, int count) { if (FailRead) throw new InvalidOperationException("Code unavailable"); return Source; }
            public void DeleteLines(int start, int count)
            {
                DeleteAttempts++; BeforeDelete?.Invoke();
                if (start != 1 || count < 1 || !Source.StartsWith(string.Concat(Enumerable.Repeat("\r\n", count)), StringComparison.Ordinal))
                    throw new InvalidOperationException("Only the proved leading CRLF prefix may be removed.");
                Source = Source.Substring(count * 2);
                if (ThrowAfterDelete) throw new InvalidOperationException("Delete applied then threw.");
            }
        }
        public sealed class Components : IEnumerable<Component>
        {
            public List<Component> Items { get; } = new List<Component>();
            public string Fault;
            public readonly Dictionary<string, string> ImportedCodeOverrides = new Dictionary<string, string>(StringComparer.Ordinal);
            public readonly Dictionary<string, string> FormExportHeaderOverrides = new Dictionary<string, string>(StringComparer.Ordinal);
            public int ImportAttempts;
            public int RemoveAttempts;
            public void Remove(Component component) { RemoveAttempts++; Items.Remove(component); }
            public Component Import(string path)
            {
                ImportAttempts++;
                string text = File.ReadAllText(path);
                string name = System.Text.RegularExpressions.Regex.Match(text, "VB_Name = \"([^\"]+)\"").Groups[1].Value;
                var component = new Component(name, System.IO.Path.GetExtension(path) == ".cls" ? 2 : System.IO.Path.GetExtension(path) == ".frm" ? 3 : 1);
                component.CodeModule.Source = ImportedCodeOverrides.ContainsKey(name) ? ImportedCodeOverrides[name] : Fault == "wrong-code" ? "Wrong" : "Option Explicit";
                if (FormExportHeaderOverrides.ContainsKey(name)) component.FormExportHeader = FormExportHeaderOverrides[name];
                Items.Add(component);
                if (Fault == "throw") throw new InvalidOperationException("Native import threw after applying.");
                return component;
            }
            public IEnumerator<Component> GetEnumerator() => Items.GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
        public sealed class Reference
        {
            public string GUID { get; set; } = "{000204EF-0000-0000-C000-000000000046}";
            public int Major { get; set; } = 4;
            public int Minor { get; set; } = 2;
            public bool IsBroken { get; set; }
            public bool BuiltIn { get; set; }
            public int Type { get; set; }
        }
        public sealed class References : IEnumerable<Reference>
        {
            public List<Reference> Items { get; } = new List<Reference>();
            public int AddAttempts;
            public Reference AddFromGuid(string guid, int major, int minor)
            {
                AddAttempts++; var reference = new Reference { GUID = guid, Major = major, Minor = minor }; Items.Add(reference); return reference;
            }
            public IEnumerator<Reference> GetEnumerator() => Items.GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
        private sealed class Owner : VbeProjectComponents.ISolidWorksSaveProbe
        {
            public bool IsSolidWorks => true;
            public int ProcessId => 1;
            public VbeProjectComponents.SolidWorksSaveSelection Selection => null;
            public void RequireOwner(object editor) { }
            public bool SameProject(object a, object b) => ReferenceEquals(a, b);
            public object SelectComponent(object editor, object project) => null;
            public void RestoreSelection(object editor, object project, object component) { }
            public bool SelectionMatches(object editor, object project, object component) => true;
            public object SaveControl(object editor) => null;
            public void Save(object control) { }
            public bool FileExists(string path) => File.Exists(path);
            public bool FileReadOnly(string path) => false;
            public long FileLength(string path) => new FileInfo(path).Length;
        }
    }
}
