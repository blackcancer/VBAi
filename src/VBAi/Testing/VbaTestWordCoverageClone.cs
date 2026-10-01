using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace VBAi
{
    /// <summary>Copies the saved Word file; never saves or converts the original document.</summary>
    internal sealed class VbaTestWordCoverageClone
    {
        internal VbaTestWordValuesHost Host = new VbaTestWordValuesHost();
        internal Action<string, string> CopySavedFile = CopyFile;
        internal Func<object, string, object> OpenCopy = (application, path) => ((dynamic)application).Documents.Open(
            FileName: path, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, Revert: false, Visible: false, OpenAndRepair: false);
        internal Action<object> CloseCopy = document => ((dynamic)document).Close(SaveChanges: 0);

        internal static VbaTestCoverageClone CreateWord(object project, string sourcePath, string folder)
        { return new VbaTestWordCoverageClone().Create(project, sourcePath, folder); }

        internal VbaTestCoverageClone Create(object project, string sourcePath, string folder)
        {
            var source = (VbaTestWordValuesHost.OwnedTarget)Host.ResolveTarget(project, sourcePath);
            RequireSaved(source.Document);
            string extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (!new[] { ".docm", ".dotm", ".doc", ".dot" }.Contains(extension))
                throw new InvalidOperationException("Word coverage requires a saved DOCM, DOTM, DOC or DOT file.");
            VbaTestWordValuesHost.RequireAbsolutePath(folder);
            string copyPath = Path.Combine(Path.GetFullPath(folder), "coverage" + extension);
            if (VbaTestWordValuesHost.SamePath(copyPath, sourcePath) || File.Exists(copyPath))
                throw new InvalidOperationException("The owned Word copy path is already occupied.");
            if (Directory.Exists(folder)) throw new InvalidOperationException("Word coverage requires a new unique output directory.");
            foreach (dynamic document in ((dynamic)source.Application).Documents)
                if (VbaTestWordValuesHost.SamePath((string)document.FullName, copyPath))
                    throw new InvalidOperationException("The Word copy path already belongs to an open document.");
            Host.ValidateTarget(source);
            RequireSaved(source.Document);
            Directory.CreateDirectory(folder);
            // Word retains a writer handle: share it, then protect and verify the saved bytes during copying.
            // Unsaved VBA changes are rejected by the service's full source/reference comparison.
            try { CopySavedFile(source.Path, copyPath); }
            catch (Exception error) { throw Uncertain("Copying the saved Word file", copyPath, error); }
            try { Host.ValidateTarget(source); RequireSaved(source.Document); }
            catch (Exception error) { throw Uncertain("Verifying the original Word document after copying", copyPath, error); }
            object copy;
            // Word AutoOpen/Document_Open and application events may run here. No security,
            // trust, EnableEvents or DisableAutoMacros settings are changed by this provider.
            try { copy = OpenCopy(source.Application, copyPath); }
            catch (Exception error) { throw Uncertain("Opening the Word copy", copyPath, error); }
            try
            {
                if (copy == null || Host.SameIdentity(copy, source.Document)
                    || Host.SameIdentity((object)((dynamic)copy).VBProject, source.Project)
                    || !VbaTestWordValuesHost.SamePath((string)((dynamic)copy).FullName, copyPath))
                    throw new InvalidOperationException("The opened Word coverage document is not a verified distinct owned copy. It was not closed.");
                var ownedCopy = (VbaTestWordValuesHost.OwnedTarget)Host.ResolveTarget((object)((dynamic)copy).VBProject, copyPath);
                if (!Host.SameIdentity(ownedCopy.Document, copy)) throw new InvalidOperationException("The Word copy identity could not be verified. It was not closed.");
                Host.ValidateTarget(source);
                RequireSaved(source.Document);
                return new VbaTestCoverageClone { Project = ownedCopy.Project, Path = copyPath, Close = () =>
                {
                    var current = Host.ValidateTarget(ownedCopy);
                    if (Host.SameIdentity(current.Document, source.Document) || Host.SameIdentity(current.Project, source.Project))
                        throw new InvalidOperationException("Closing the original Word document is forbidden.");
                    try
                    {
                        CloseCopy(current.Document);
                        Host.RequireOwner();
                        object application = Host.ResolveApplication();
                        if (!Host.SameIdentity(application, current.Application)) throw new InvalidOperationException("The Word application identity changed during close.");
                        int count = 0;
                        foreach (dynamic document in ((dynamic)application).Documents)
                        {
                            if (++count > 1000) throw new InvalidOperationException("Unexpected Word document count after close.");
                            if (Host.SameIdentity((object)document, current.Document))
                                throw new InvalidOperationException("The owned Word copy remains open; close may have been cancelled.");
                        }
                    }
                    catch (Exception error) { throw Uncertain("Closing the owned Word copy", copyPath, error); }
                } };
            }
            catch (Exception error)
            { throw Uncertain("Verifying ownership after opening the Word copy", copyPath, error); }
        }

        private static void CopyFile(string source, string destination)
        { CopyFile(source, destination, null); }

        internal static void CopyFile(string source, string destination, Action<FileStream, FileStream> afterCopy)
        {
            // Share existing Word writers without allowing deletion or replacement of this source path.
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Lock beyond EOF as well, so a competing WriteFile cannot append during the snapshot.
                input.Lock(0, long.MaxValue);
                try
                {
                    long length = input.Length;
                    DateTime modified = File.GetLastWriteTimeUtc(source);
                    using (var hash = SHA256.Create())
                    {
                        byte[] expected = hash.ComputeHash(input);
                        input.Position = 0;
                        using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                        {
                            input.CopyTo(output);
                            output.Flush(true);
                            afterCopy?.Invoke(input, output);
                            if (input.Length != length || output.Length != length || File.GetLastWriteTimeUtc(source) != modified)
                                throw new IOException("The saved Word file metadata changed during copying; the destination was retained.");
                            // Byte-range locks do not block mapped views; verify both source passes and copied bytes.
                            input.Position = 0;
                            byte[] current = hash.ComputeHash(input);
                            output.Position = 0;
                            byte[] copied = hash.ComputeHash(output);
                            if (!expected.SequenceEqual(current) || !expected.SequenceEqual(copied))
                                throw new IOException("The saved Word bytes changed during copying; the destination was retained.");
                        }
                    }
                }
                finally { input.Unlock(0, long.MaxValue); }
            }
        }
        private static void RequireSaved(object document)
        {
            object saved = ((dynamic)document).Saved;
            if (!(saved is bool value) || !value)
                throw new InvalidOperationException("Save the original Word document before collecting coverage; unsaved document content cannot be reproduced by a file copy.");
        }
        private static VbaTestInvocationException Uncertain(string phase, string retainedPath, Exception error)
        { return new VbaTestInvocationException(phase + " did not return verified completion; no retry was attempted. Retained copy: " + retainedPath + ". " + error.Message, true, error); }
    }
}
