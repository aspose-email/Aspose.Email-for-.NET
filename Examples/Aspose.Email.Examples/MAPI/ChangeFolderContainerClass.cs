// Demonstrates how to change the container class of a PST folder (e.g., to IPF.Note).

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ChangeFolderContainerClass
    {
        public static void Run()
        {
            // Work on a copy: examples must never modify the shared input data.
            var pstPath = Data.Out/"ChangeFolderContainerClass_out.pst";
            File.Copy(Data.Mapi/"PersonalStorage1.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var folder = pst.RootFolder.GetSubFolder("Inbox");

                // The container class tells Outlook what kind of items a folder holds.
                // This one starts out as IPF.Note (a mail folder); marking it IPF.Imap
                // makes Outlook treat it as an IMAP folder instead.
                Console.WriteLine($"Container class before: {folder.ContainerClass}");
                folder.ChangeContainerClass("IPF.Imap");
                Console.WriteLine($"Container class after:  {folder.ContainerClass}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}
