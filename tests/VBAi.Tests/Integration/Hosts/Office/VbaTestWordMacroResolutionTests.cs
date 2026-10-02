using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("VbaTestWordMacroResolution"), DoNotParallelize]
    public sealed class VbaTestWordMacroResolutionTests
    {
        [STATestMethod]
        [DataRow(0)]
        [DataRow(2)]
        public void UniqueOwnedModuleReturnsBooleanAndBooleanArray(int argumentCount)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Native Word diagnostics require VBAi_RUN_OFFICE_TESTS=1.");
            using (var fixture = OfficeVbeFixture.Start("Word", allowExistingHost: true, allowForcedTermination: false))
            {
                ExecuteOwnedProbe(fixture, argumentCount);
                // Only finalize unreachable probe RCWs after a confirmed terminal native result.
                if (!fixture.NativeExecutionUnsettled) {
                    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                }
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ExecuteOwnedProbe(OfficeVbeFixture fixture, int argumentCount)
        {
                const string module = "VBAiWordResolution";
                fixture.Data("create_module", "Module", module, "ExpectedMode", 2);
                var before = fixture.Data("read_module", "Module", module);
                fixture.Data("replace_lines", "Module", module, "StartLine", 1, "Count", 0,
                    "ExpectedSha256", before["Sha256"], "Text", "Option Explicit\r\nPublic Function ReadZero(Optional ByVal first As Variant, Optional ByVal second As Variant) As Boolean\r\nReadZero = True\r\nActiveDocument.Variables.Add \"VBAiWordRunValue\", \"owned\"\r\nEnd Function\r\nPublic Function ReadTwo(ByVal first As String, ByVal second As String) As Variant\r\nDim hits(1 To 2) As Boolean\r\nhits(1) = True\r\nReadTwo = hits\r\nActiveDocument.Variables.Add \"VBAiWordRunValue\", first & second\r\nEnd Function\r\n");
                fixture.SaveNative();
                Assert.AreEqual(1, fixture.Items("list_projects").Count(project => Equals(project["Name"], fixture.Project)),
                    "The project qualifier must be unique across the loaded VBE projects.");
                object application = Marshal.GetActiveObject("Word.Application");
                object document = null, projectObject = null, variables = null, markerVariable = null;
                int? originalSecurity = null;
                VbaTestWordValuesHost.OwnedTarget target = null;
                try
                {
                    document = ((dynamic)application).ActiveDocument;
                    projectObject = ((dynamic)document).VBProject;
                    var host = new VbaTestWordValuesHost { ReadProcessName = () => "WINWORD", ReadProcessId = () => fixture.ProcessId,
                        ReadActiveApplication = unused => application };
                    uint owner = host.ReadWindowOwner(VbaTestWordValuesHost.ReadApplicationWindow(application));
                    Assert.AreEqual((uint)fixture.ProcessId, owner);
                    Assert.AreEqual(fixture.DocumentPath, (string)((dynamic)document).FullName, true);
                    Assert.AreEqual(fixture.Project, (string)((dynamic)projectObject).Name);
                    originalSecurity = Convert.ToInt32(((dynamic)application).AutomationSecurity);
                    ((dynamic)application).AutomationSecurity = 2;
                    Assert.AreEqual(2, Convert.ToInt32(((dynamic)application).AutomationSecurity));

                    target = (VbaTestWordValuesHost.OwnedTarget)host.ResolveTarget(projectObject, fixture.DocumentPath);
                    string procedure = argumentCount == 0 ? "ReadZero" : "ReadTwo";
                    int matchingModules = 0;
                    foreach (var loaded in fixture.Items("list_projects"))
                    {
                        var modules = fixture.Response("list_modules", "Project", loaded["Name"]);
                        Assert.AreEqual(true, modules["Ok"]);
                        matchingModules += ((object[])modules["Data"]).Select(VbeBridgeClient.Object)
                            .Count(component => Equals(component["Name"], module));
                    }
                    Assert.AreEqual(1, matchingModules, "Only the exact owned project may expose this synthetic module.");
                    string qualifier = module + "." + procedure;
                    fixture.NativeExecutionUnsettled = true;
                    object result = host.Invoke(target, module, procedure, argumentCount == 0 ? new object[] { false, false } : new object[] { "own", "ed" });
                    variables = ((dynamic)document).Variables;
                    markerVariable = ((dynamic)variables)["VBAiWordRunValue"];
                    string marker = Convert.ToString(((dynamic)markerVariable).Value);
                    fixture.NativeExecutionUnsettled = false;
                    Assert.AreEqual("owned", marker, "The exact synthetic function must actually execute.");
                    if (argumentCount == 0) Assert.AreEqual(true, result);
                    else {
                        var hits = result as Array;
                        Assert.IsNotNull(hits, "The VBA Boolean SAFEARRAY must cross the native Word boundary.");
                        CollectionAssert.AreEqual(new[] { true, false }, hits.Cast<bool>().ToArray());
                    }
                    File.WriteAllText(Path.Combine(fixture.Root, "word-name-resolution.json"), new JavaScriptSerializer().Serialize(new {
                        fixture.ProcessId, fixture.DocumentPath, Qualifier = qualifier, NativeArguments = 2, OptionalFunction = argumentCount == 0,
                        Result = result, Marker = marker, DiagnosticOnly = true, NativeCalls = 1 }));
                }
                catch (Exception error)
                {
                    File.WriteAllText(Path.Combine(fixture.Root, "word-name-resolution-error.json"), error.ToString());
                    throw;
                }
                finally
                {
                    try {
                        if (originalSecurity.HasValue && !fixture.NativeExecutionUnsettled)
                            ((dynamic)application).AutomationSecurity = originalSecurity.Value;
                    }
                    finally {
                        target?.Dispose();
                        var context = new[] { markerVariable, variables, projectObject, document, application };
                        if (fixture.NativeExecutionUnsettled) VbaTestWordValuesHost.RetainAcquired(context);
                        else foreach (var item in context)
                            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
                    }
                }
        }
    }
}
