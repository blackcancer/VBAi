using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{
    // This generator never modifies a live project. Its explicit dispatch table is
    // installed through a separate, reviewed project edit before any run.
    internal static class VbaTestRuntimeSource
    {
        internal const string ModuleName = "VBAiTestSupport";
        internal const string Version = "3";
        internal const string DispatcherProcedure = "VBAiExecuteTest";
        internal const string PendingProcedure = "VBAiExecutePendingTest";
        internal const int MaximumMessageLength = 8192;
        internal const int MaximumCasesPerLeaf = 128;
        internal const int MaximumLeafBodyCharacters = 16 * 1024;
        internal const int MaximumRouteChildren = 32;

        private const string Header = "Option Explicit\n' VBAi test support version " + Version
            + "\n' Generated dispatch only; refresh this module explicitly after changing tests.\n";

        private static readonly Regex Identifier = new Regex(@"\A\p{L}[\p{L}\p{N}_]{0,254}\z", RegexOptions.CultureInvariant);

        internal static string Generate(VbaTestCatalog catalog)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            var supportModules = (catalog.Project?.Modules ?? new VbaTestModuleSnapshot[0])
                .Where(x => string.Equals(x.Name, ModuleName, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (supportModules.Length > 1 || (supportModules.Length == 1 && !IsOwned(supportModules[0].Source)))
                throw new InvalidOperationException("An existing module conflicts with the VBAi test support module name.");
            var candidates = new List<VbaTestDescriptor>();
            var modules = catalog.Modules ?? new List<VbaTestModule>();
            var duplicateModules = new HashSet<string>(modules.GroupBy(x => x.Name ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Where(x => x.Count() > 1).Select(x => x.Key), StringComparer.OrdinalIgnoreCase);
            foreach (var module in modules)
            {
                if (!SafeIdentifier(module.Name) || duplicateModules.Contains(module.Name)
                    || string.Equals(module.Name, ModuleName, StringComparison.OrdinalIgnoreCase)
                    || !string.IsNullOrEmpty(module.Diagnostic)) continue;
                candidates.AddRange((module.Tests ?? new List<VbaTestDescriptor>()).Where(x => Eligible(x, module.Name)));
                foreach (var fixture in new[] { module.ModuleInitialize, module.ModuleCleanup, module.TestInitialize, module.TestCleanup })
                    if (Eligible(fixture, module.Name)) candidates.Add(fixture);
            }

            // An ambiguous call cannot be repaired by arbitrarily choosing a descriptor.
            var unique = candidates.GroupBy(x => x.Module + "." + x.Procedure, StringComparer.OrdinalIgnoreCase)
                .Where(x => x.Count() == 1).Select(x => x.Single())
                .OrderBy(x => x.Module, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.Procedure, StringComparer.OrdinalIgnoreCase).ToArray();
            var helpers = new StringBuilder();
            var leaves = new List<string>();
            var body = new StringBuilder();
            int cases = 0;
            foreach (var test in unique)
            {
                string branch = DispatchBranch(test);
                if (cases > 0 && (cases == MaximumCasesPerLeaf || body.Length + branch.Length > MaximumLeafBodyCharacters))
                { AppendLeaf(helpers, leaves, body); cases = 0; }
                body.Append(branch); cases++;
            }
            // An empty catalogue still has an explicit unknown-key path and no executable test cases.
            if (cases > 0 || leaves.Count == 0) AppendLeaf(helpers, leaves, body);
            string root = AppendRoutes(helpers, leaves);
            var source = new StringBuilder();
            source.Append(Header.Replace("\n", "\r\n"));
            source.AppendLine(RuntimePrefix.Replace("\r\n", "\n").Replace("\n", "\r\n"));
            source.Append("    If " + root + "(LCase$(moduleName & \".\" & procedureName)) Then GoTo Executed\r\n"
                + "    mRunning = False\r\n"
                + "    VBAiExecuteTest = Array(\"Error\", \"This procedure is absent from the installed test dispatch table.\", \"0\")\r\n"
                + "    Exit Function\r\n");
            source.Append(RuntimeSuffix.Replace("\r\n", "\n").Replace("\n", "\r\n"));
            source.Append(helpers);
            source.Append("\r\nPublic Sub VBAiExecutePendingTest()\r\n    Dim channel As Object\r\n    Dim job As Variant\r\n    Dim result As Variant\r\n"
                + "    Set channel = CreateObject(\"VBAi.TestRuntime\")\r\n"
                + "    job = channel.Request(\"" + Version + "\", \"" + DispatchSignature(catalog) + "\")\r\n"
                + "    result = VBAiExecuteTest(CStr(job(1)), CStr(job(2)))\r\n"
                + "    If Not channel.Publish(CStr(job(0)), CStr(job(4)), CStr(job(5)), CStr(job(6)), CStr(result(0)), CStr(result(1)), CLng(result(2))) Then\r\n"
                + "        Err.Raise vbObjectError + 2050, \"VBAi.TestRuntime\", \"The test result was not accepted.\"\r\n"
                + "    End If\r\nEnd Sub\r\n");
            return source.ToString();
        }

        private static string DispatchBranch(VbaTestDescriptor test)
        {
            string target = test.Module + "." + test.Procedure;
            var source = new StringBuilder("        Case LCase$(\"" + target + "\")\r\n");
            if (test.Kind == "Function")
                source.Append("            If Not " + target + "() Then\r\n"
                    + "                If Not mFailed Then mMessage = \"Boolean test returned False.\"\r\n"
                    + "                mFailed = True\r\n            End If\r\n");
            else source.Append("            Call " + target + "\r\n");
            return source.ToString();
        }

        private static void AppendLeaf(StringBuilder source, List<string> names, StringBuilder body)
        {
            string name = "VBAiDispatchLeaf" + names.Count.ToString("D6", CultureInfo.InvariantCulture);
            names.Add(name);
            source.Append("\r\nPrivate Function " + name + "(ByVal key As String) As Boolean\r\n    Select Case key\r\n");
            source.Append(body);
            source.Append("        Case Else\r\n            Exit Function\r\n    End Select\r\n"
                + "    " + name + " = True\r\nEnd Function\r\n");
            body.Clear();
        }

        private static string AppendRoutes(StringBuilder source, List<string> children)
        {
            int route = 0;
            while (children.Count > 1)
            {
                var parents = new List<string>();
                for (int offset = 0; offset < children.Count; offset += MaximumRouteChildren)
                {
                    string name = "VBAiDispatchRoute" + (route++).ToString("D6", CultureInfo.InvariantCulture);
                    parents.Add(name);
                    source.Append("\r\nPrivate Function " + name + "(ByVal key As String) As Boolean\r\n");
                    // VBA Or is not short-circuiting. Explicit early returns prevent multiple native invocations.
                    foreach (string child in children.Skip(offset).Take(MaximumRouteChildren))
                        source.Append("    If " + child + "(key) Then\r\n        " + name
                            + " = True\r\n        Exit Function\r\n    End If\r\n");
                    source.Append("End Function\r\n");
                }
                children = parents;
            }
            return children[0];
        }

        internal static string DispatchSignature(VbaTestCatalog catalog)
        {
            var entries = catalog.Modules.SelectMany(module => (module.Tests ?? new List<VbaTestDescriptor>()).Concat(
                new[] { module.ModuleInitialize, module.ModuleCleanup, module.TestInitialize, module.TestCleanup }.Where(item => item != null)))
                .Select(item => (item.Module + "." + item.Procedure + "|" + item.Kind + "|" + item.Diagnostic + "|" + item.IgnoreReason).ToLowerInvariant())
                .OrderBy(item => item, StringComparer.Ordinal).ToArray();
            return VbeTestExplorerService.Hash(string.Join("\n", entries));
        }

        // Ownership is a versioned marker, not proof that the source is unchanged.
        // Installation must also back up the old source and review the replacement.
        internal static bool IsOwned(string source)
        {
            if (source == null) return false;
            string canonical = source.Replace("\r\n", "\n");
            return canonical.StartsWith(Header, StringComparison.Ordinal) || canonical.StartsWith(
                "Option Explicit\n' VBAi test support version 1\n' Generated dispatch only; refresh this module explicitly after changing tests.\n", StringComparison.Ordinal) || canonical.StartsWith(
                "Option Explicit\n' VBAi test support version 2\n' Generated dispatch only; refresh this module explicitly after changing tests.\n", StringComparison.Ordinal);
        }

        internal static VbaTestResult Decode(VbaTestDescriptor test, object native)
        {
            if (test == null) throw new ArgumentNullException(nameof(test));
            var array = native as Array;
            if (array == null || array.Rank != 1 || array.Length != 3)
                throw new InvalidOperationException("The VBA test did not return a three-field result envelope.");
            var offset = array.GetLowerBound(0);
            var status = array.GetValue(offset) as string;
            var message = array.GetValue(offset + 1) as string;
            var errorText = array.GetValue(offset + 2) as string;
            VbaTestOutcome outcome;
            int errorNumber;
            if (status == null || status.Length > 32 || !Enum.TryParse(status, false, out outcome)
                || (outcome != VbaTestOutcome.Passed && outcome != VbaTestOutcome.Failed
                    && outcome != VbaTestOutcome.Error && outcome != VbaTestOutcome.Inconclusive)
                || !string.Equals(status, outcome.ToString(), StringComparison.Ordinal)
                || message == null || message.Length > MaximumMessageLength
                || errorText == null || errorText.Length > 12
                || !int.TryParse(errorText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out errorNumber)
                || ((outcome == VbaTestOutcome.Passed || outcome == VbaTestOutcome.Inconclusive) && errorNumber != 0))
                throw new InvalidOperationException("The VBA test returned an invalid result envelope.");
            return new VbaTestResult { Test = test, Outcome = outcome, Message = message, ErrorNumber = errorNumber };
        }

        private static bool SafeIdentifier(string name) => name != null && Identifier.IsMatch(name);

        private static bool Eligible(VbaTestDescriptor test, string module)
        {
            return test != null && string.Equals(test.Module, module, StringComparison.OrdinalIgnoreCase)
                && SafeIdentifier(test.Procedure) && string.IsNullOrEmpty(test.Diagnostic)
                && string.IsNullOrEmpty(test.IgnoreReason)
                && (test.Kind == "Sub" || test.Kind == "Function");
        }

        private const string RuntimePrefix = @"
Private Const FailureError As Long = vbObjectError + 2048
Private Const InconclusiveError As Long = vbObjectError + 2049
Private mFailed As Boolean
Private mInconclusive As Boolean
Private mRunning As Boolean
Private mMessage As String

Public Function VBAiExecuteTest(ByVal moduleName As String, ByVal procedureName As String) As Variant
    Dim outcome As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String
    If mRunning Then
        VBAiExecuteTest = Array(""Error"", ""A VBA test is already running."", ""0"")
        Exit Function
    End If
    mRunning = True
    mFailed = False
    mInconclusive = False
    mMessage = vbNullString
    On Error GoTo FailedCall
";

        private const string RuntimeSuffix = @"Executed:
    If mFailed Then
        outcome = ""Failed""
    ElseIf mInconclusive Then
        outcome = ""Inconclusive""
    Else
        outcome = ""Passed""
    End If
    GoTo Publish
FailedCall:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
    Err.Clear
    If errorNumber = FailureError And mFailed Then
        outcome = ""Failed""
    ElseIf errorNumber = InconclusiveError And mInconclusive Then
        If mFailed Then
            outcome = ""Failed""
        Else
            outcome = ""Inconclusive""
            errorNumber = 0
        End If
    Else
        outcome = ""Error""
        mMessage = errorDescription
        If Len(errorSource) > 0 Then mMessage = errorSource & "": "" & mMessage
    End If
Publish:
    mRunning = False
    VBAiExecuteTest = Array(outcome, Left$(mMessage, 8192), CStr(errorNumber))
End Function

Public Sub Fail(Optional ByVal message As String = ""Assertion failed."")
    If Not mRunning Then Err.Raise 5, ""VBAiTestSupport"", ""Assertions require an active test run.""
    If Not mFailed Then mMessage = Left$(message, 8192)
    mFailed = True
    Err.Raise FailureError, ""VBAiTestSupport"", Left$(message, 8192)
End Sub

Public Sub Inconclusive(Optional ByVal message As String = ""Test is not implemented."")
    If Not mRunning Then Err.Raise 5, ""VBAiTestSupport"", ""Assertions require an active test run.""
    If Not mFailed And Not mInconclusive Then mMessage = Left$(message, 8192)
    mInconclusive = True
    Err.Raise InconclusiveError, ""VBAiTestSupport"", Left$(message, 8192)
End Sub

Public Sub IsTrue(ByVal actual As Variant, Optional ByVal message As String = ""Expected True."")
    If VarType(actual) <> vbBoolean Then
        Fail ""IsTrue requires a Boolean value.""
        Exit Sub
    End If
    If Not actual Then Fail message
End Sub

Public Sub IsFalse(ByVal actual As Variant, Optional ByVal message As String = ""Expected False."")
    If VarType(actual) <> vbBoolean Then
        Fail ""IsFalse requires a Boolean value.""
        Exit Sub
    End If
    If actual Then Fail message
End Sub

Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, _
                    Optional ByVal message As String = ""Values are not equal."")
    Dim scalarType As Integer
    If IsObject(expected) Or IsObject(actual) Then
        Fail ""AreEqual does not support objects.""
        Exit Sub
    End If
    If IsArray(expected) Or IsArray(actual) Then
        Fail ""AreEqual does not support arrays.""
        Exit Sub
    End If
    scalarType = VarType(expected)
    If scalarType <> VarType(actual) Then
        Fail ""AreEqual requires identical scalar types; convert explicitly.""
        Exit Sub
    End If
    Select Case scalarType
        Case vbString
            If StrComp(expected, actual, vbBinaryCompare) <> 0 Then Fail message
        Case vbBoolean, vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal, 20
            If expected <> actual Then Fail message
        Case Else
            Fail ""AreEqual does not support Empty, Null, Error or this value type.""
    End Select
End Sub
";
    }
}
