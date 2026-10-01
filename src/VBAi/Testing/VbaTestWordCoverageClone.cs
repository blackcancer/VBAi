using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;

namespace VBAi
{
    /// <summary>Copies the saved Word file; never saves or converts the original document.</summary>
    internal sealed class VbaTestWordCoverageClone
    {
        internal VbaTestWordValuesHost Host = new VbaTestWordValuesHost();
        internal Action<string, string> CopySavedFile = CopyFile;
        internal Func<object, string, object> OpenCopy = OpenDocument;
        internal Action<object> CloseCopy = document => ((dynamic)document).Close(SaveChanges: 0);

        internal static VbaTestCoverageClone CreateWord(object project, string sourcePath, string folder)
        { return new VbaTestWordCoverageClone().Create(project, sourcePath, folder); }

        internal VbaTestCoverageClone Create(object project, string sourcePath, string folder)
        {
            var references = new CloneReferences();
            bool transferred = false;
            try
            {
                var source = references.Source = (VbaTestWordValuesHost.OwnedTarget)Host.ResolveTarget(project, sourcePath);
                RequireSaved(source.Document);
                string extension = Path.GetExtension(sourcePath).ToLowerInvariant();
                if (!new[] { ".docm", ".dotm", ".doc", ".dot" }.Contains(extension))
                    throw new InvalidOperationException("Word coverage requires a saved DOCM, DOTM, DOC or DOT file.");
                VbaTestWordValuesHost.RequireAbsolutePath(folder);
                string copyPath = Path.Combine(Path.GetFullPath(folder), "coverage" + extension);
                if (VbaTestWordValuesHost.SamePath(copyPath, sourcePath) || File.Exists(copyPath))
                    throw new InvalidOperationException("The owned Word copy path is already occupied.");
                if (Directory.Exists(folder)) throw new InvalidOperationException("Word coverage requires a new unique output directory.");
                InspectDocuments(source.Application, false, document => {
                    if (VbaTestWordValuesHost.SamePath((string)((dynamic)document).FullName, copyPath))
                        throw new InvalidOperationException("The Word copy path already belongs to an open document.");
                });
                Host.ValidateTarget(source);
                RequireSaved(source.Document);
                Directory.CreateDirectory(folder);
                // Word retains a writer handle: share it, then protect and verify the saved bytes during copying.
                try { CopySavedFile(source.Path, copyPath); }
                catch (Exception error) { throw Uncertain("Copying the saved Word file", copyPath, error); }
                try { Host.ValidateTarget(source); RequireSaved(source.Document); }
                catch (Exception error) { throw Uncertain("Verifying the original Word document after copying", copyPath, error); }
                // Word document events may execute here; host security and trust remain unchanged.
                try { references.OpenedDocument = OpenCopy(source.Application, copyPath); }
                catch (Exception error) { throw Uncertain("Opening the Word copy", copyPath, error); }
                try
                {
                    object copy = references.OpenedDocument;
                    if (copy == null || Host.SameIdentity(copy, source.Document))
                        throw new InvalidOperationException("The opened Word coverage document is not a verified distinct owned copy. It was not closed.");
                    references.CopyProject = ((dynamic)copy).VBProject;
                    if (Host.SameIdentity(references.CopyProject, source.Project)
                        || !VbaTestWordValuesHost.SamePath((string)((dynamic)copy).FullName, copyPath))
                        throw new InvalidOperationException("The opened Word coverage document is not a verified distinct owned copy. It was not closed.");
                    var ownedCopy = references.Copy = (VbaTestWordValuesHost.OwnedTarget)Host.ResolveTarget(references.CopyProject, copyPath);
                    if (!Host.SameIdentity(ownedCopy.Document, copy)) throw new InvalidOperationException("The Word copy identity could not be verified. It was not closed.");
                    Host.ValidateTarget(source);
                    RequireSaved(source.Document);
                    var clone = new VbaTestCoverageClone { Project = ownedCopy.Project, Path = copyPath, Close = () => CloseOwnedCopy(references, copyPath) };
                    transferred = true;
                    return clone;
                }
                catch (Exception error) { throw Uncertain("Verifying ownership after opening the Word copy", copyPath, error); }
            }
            catch (VbaTestInvocationException error)
            {
                if (error.Uncertain) references.Retain();
                throw;
            }
            finally { if (!transferred) references.Dispose(); }
        }

        private void CloseOwnedCopy(CloneReferences references, string copyPath)
        {
            try
            {
                var current = Host.ValidateTarget(references.Copy);
                if (Host.SameIdentity(current.Document, references.Source.Document) || Host.SameIdentity(current.Project, references.Source.Project))
                    throw new InvalidOperationException("Closing the original Word document is forbidden.");
                try
                {
                    CloseCopy(current.Document);
                    Host.RequireOwner();
                    using (var application = Host.ResolveApplicationLease())
                    {
                        try
                        {
                            if (!Host.SameIdentity(application.Application, current.Application)) throw new InvalidOperationException("The Word application identity changed during close.");
                            InspectDocuments(application.Application, true, document => {
                                if (Host.SameIdentity(document, current.Document))
                                    throw new InvalidOperationException("The owned Word copy remains open; close may have been cancelled.");
                            });
                        }
                        catch { application.RetainOnUncertain(); throw; }
                    }
                }
                catch (Exception error) { throw Uncertain("Closing the owned Word copy", copyPath, error); }
                // The copied project remains leased until both Close and absence are verified.
                references.Dispose();
            }
            catch { references.Retain(); throw; }
        }

        private void InspectDocuments(object application, bool retainOnFailure, Action<object> inspect)
        {
            var acquired = new List<object>();
            bool failed = false;
            try
            {
                object documents = ((dynamic)application).Documents;
                acquired.Add(documents);
                int count = (int)((dynamic)documents).Count;
                if (count < 0 || count > 1000) throw new InvalidOperationException("Unexpected Word document count.");
                for (int index = 1; index <= count; index++)
                {
                    object document = Host.ReadDocumentItem(documents, index);
                    acquired.Add(document);
                    inspect(document);
                }
            }
            catch { failed = retainOnFailure; throw; }
            finally
            {
                try
                {
                    for (int index = acquired.Count - 1; index >= 0; index--)
                        if (failed) VbaTestWordValuesHost.RetainAcquired(acquired[index]);
                        else VbaTestWordValuesHost.ReleaseAcquired(acquired[index]);
                }
                catch { VbaTestWordValuesHost.RetainAcquired(acquired); throw; }
            }
        }

        private static object OpenDocument(object application, string path)
        {
            object documents = ((dynamic)application).Documents;
            bool uncertain = false;
            object opened = null;
            try
            {
                opened = ((dynamic)documents).Open(FileName: path, ConfirmConversions: false, ReadOnly: false,
                    AddToRecentFiles: false, Revert: false, Visible: false, OpenAndRepair: false);
                return opened;
            }
            catch { uncertain = true; throw; }
            finally
            {
                if (uncertain) VbaTestWordValuesHost.RetainAcquired(documents);
                else
                {
                    try { VbaTestWordValuesHost.ReleaseAcquired(documents); }
                    catch { VbaTestWordValuesHost.RetainAcquired(opened); VbaTestWordValuesHost.RetainAcquired(documents); throw; }
                }
            }
        }

        private sealed class CloneReferences : IDisposable
        {
            internal VbaTestWordValuesHost.OwnedTarget Source, Copy;
            internal object OpenedDocument, CopyProject;
            private bool retained, disposed;
            internal void Retain()
            {
                if (retained) return;
                retained = true;
                Source.RetainOnUncertain();
                Copy?.RetainOnUncertain();
                VbaTestWordValuesHost.RetainAcquired(this);
            }
            private static void ReleaseOnce(ref object reference)
            {
                object acquired = reference;
                reference = null;
                try { VbaTestWordValuesHost.ReleaseAcquired(acquired); }
                catch { VbaTestWordValuesHost.RetainAcquired(acquired); throw; }
            }
            public void Dispose()
            {
                if (retained || disposed) return;
                disposed = true;
                try { Copy?.Dispose(); }
                finally
                {
                    try { ReleaseOnce(ref CopyProject); }
                    finally
                    {
                        try { ReleaseOnce(ref OpenedDocument); }
                        finally { Source?.Dispose(); }
                    }
                }
                Copy = null; Source = null; CopyProject = null; OpenedDocument = null;
            }
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
