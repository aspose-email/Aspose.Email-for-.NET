// Demonstrates how to create a PST file and add a document file to a subfolder.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddFilesToPst
    {
        public static void Run()
        {
            var path = Data.Out/"Ps1_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                var folder = pst.RootFolder.AddSubFolder("Files");

                // AddFile stores an arbitrary document in the PST, wrapped in a message.
                folder.AddFile(Data.Mapi/"attachment_1.doc", null);

                Console.WriteLine($"Added {folder.ContentCount} file(s) to the \"{folder.DisplayName}\" folder.");
            }

            Console.WriteLine($"Saved to {path}");
        }
    }
}
