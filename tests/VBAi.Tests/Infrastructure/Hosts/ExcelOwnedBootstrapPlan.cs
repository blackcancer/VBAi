using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using Microsoft.Win32;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure preflight and macro-free seed construction for the two owned scalar diagnostics.</summary>
    internal static class ExcelOwnedBootstrapPlan
    {
        internal static string ResolveExecutable()
        {
            using (var registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (var key = registry.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"))
            {
                string path = Convert.ToString(key?.GetValue(null));
                if (string.IsNullOrWhiteSpace(path)) throw new InvalidOperationException("Installed x64 Excel App Paths entry is unavailable; no activation fallback is permitted.");
                return ValidateExecutable(path);
            }
        }

        internal static string ValidateExecutable(string path)
        {
            string exact = RequireLocalAbsolutePath(path);
            if (!string.Equals(Path.GetFileName(exact), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase) || !System.IO.File.Exists(exact))
                throw new InvalidOperationException("The explicitly selected Excel executable is unavailable or has a different name.");
            using (var file = System.IO.File.OpenRead(exact))
            using (var reader = new BinaryReader(file))
            {
                if (file.Length < 64 || reader.ReadUInt16() != 0x5A4D) throw new InvalidOperationException("Excel is not a PE executable.");
                file.Position = 0x3C;
                int pe = reader.ReadInt32();
                if (pe < 64 || pe > file.Length - 6) throw new InvalidOperationException("Excel has an invalid PE header.");
                file.Position = pe;
                if (reader.ReadUInt32() != 0x00004550 || reader.ReadUInt16() != 0x8664)
                    throw new InvalidOperationException("Only the installed x64 Excel executable is selected for qualification.");
            }
            var version = FileVersionInfo.GetVersionInfo(exact);
            if (string.IsNullOrWhiteSpace(version.FileVersion) || !string.Equals(version.OriginalFilename, "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Excel executable version/original-name identity is unavailable.");
            return exact;
        }

        internal static string RequireLocalAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 3 || !char.IsLetter(path[0]) || path[1] != ':' ||
                (path[2] != '\\' && path[2] != '/') || path.IndexOf(':', 2) >= 0)
                throw new ArgumentException("An absolute local path without an alternate data stream is required.");
            string exact = Path.GetFullPath(path);
            var drive = new DriveInfo(Path.GetPathRoot(exact));
            if (drive.DriveType != DriveType.Fixed && drive.DriveType != DriveType.Removable && drive.DriveType != DriveType.Ram)
                throw new ArgumentException("A local drive is required.");
            for (string ancestor = exact; !string.IsNullOrEmpty(ancestor); ancestor = Path.GetDirectoryName(ancestor))
                if ((System.IO.File.Exists(ancestor) || Directory.Exists(ancestor)) && (System.IO.File.GetAttributes(ancestor) & FileAttributes.ReparsePoint) != 0)
                    throw new ArgumentException("Bootstrap paths must not traverse filesystem links.");
            return exact;
        }

        internal static ProcessStartInfo CreateStartInfo(string executable, string seed, string trace)
        {
            executable = RequireLocalAbsolutePath(executable);
            seed = RequireLocalAbsolutePath(seed);
            trace = RequireLocalAbsolutePath(trace);
            if (!System.IO.File.Exists(seed) || !string.Equals(Path.GetExtension(seed), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The owned macro-free xlsx seed must already exist.");
            if (!Directory.Exists(Path.GetDirectoryName(trace))) throw new ArgumentException("The trace parent directory must already exist.");
            var info = new ProcessStartInfo(executable, "/x /automation \"" + seed + "\"") {
                UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(seed),
                WindowStyle = ProcessWindowStyle.Hidden
            };
            info.EnvironmentVariables[VbeInspectionTrace.EnvironmentName] = trace;
            return info;
        }

        internal static void RequireFreshLaunch(int[] existingIds)
        {
            if (existingIds == null || existingIds.Length != 0)
                throw new InvalidOperationException("An existing Excel process refuses the explicit scalar bootstrap before launch.");
        }

        internal static void VerifyAttachedIdentity(int launchedPid, string expectedPath, string expectedStartUtc,
            int applicationPid, string actualPath, string actualStartUtc)
        {
            if (launchedPid <= 0 || applicationPid != launchedPid ||
                !string.Equals(Path.GetFullPath(expectedPath), Path.GetFullPath(actualPath), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expectedStartUtc, actualStartUtc, StringComparison.Ordinal) || string.IsNullOrWhiteSpace(expectedStartUtc))
                throw new InvalidOperationException("NativeOM Application must match the launched exact PID, executable and process start identity.");
        }

        /// <summary>Creates only an empty Open XML workbook, with no VBA, formulas, links or external relationship.</summary>
        internal static void WriteSeed(string path)
        {
            path = RequireLocalAbsolutePath(path);
            using (var package = Package.Open(path, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                var workbook = new Uri("/xl/workbook.xml", UriKind.Relative);
                var sheet = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
                WritePart(package, workbook, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"VBAiProbe\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                WritePart(package, sheet, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>");
                package.CreateRelationship(workbook, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
                package.GetPart(workbook).CreateRelationship(new Uri("worksheets/sheet1.xml", UriKind.Relative), TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId1");
            }
        }

        private static void WritePart(Package package, Uri uri, string contentType, string xml)
        {
            var part = package.CreatePart(uri, contentType, CompressionOption.Normal);
            using (var output = XmlWriter.Create(part.GetStream(FileMode.Create, FileAccess.Write), new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) }))
            using (var input = XmlReader.Create(new StringReader(xml))) output.WriteNode(input, false);
        }
    }
}
