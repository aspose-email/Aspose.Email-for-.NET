// Demonstrates how to create a new PST file and add a subfolder to its root.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateNewPstFileAndAddingSubfolders
    {
        public static void Run()
        {
            var path = Data.Out/"CreateNewPSTFileAndAddingSubfolders_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            // Unicode is the modern PST format; FileFormatVersion.Ansi would create the
            // older format with its 2 GB size limit.
            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                var inbox = pst.RootFolder.AddSubFolder("Inbox");

                Console.WriteLine($"Created {pst.Format} PST with folder \"{inbox.DisplayName}\".");
                Console.WriteLine($"Saved to {path}");
            }
        }
    }
}
