using System;
using System.IO;

namespace CodexVBE
{
    internal static class LoadLog
    {
        internal static readonly string PathName = Path.Combine(Path.GetTempPath(), "CodexVBE-load.log");

        internal static void Write(string message)
        {
            try
            {
                File.AppendAllText(PathName, DateTime.Now.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
