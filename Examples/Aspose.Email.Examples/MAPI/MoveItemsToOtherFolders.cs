// Demonstrates how to move a subfolder, a single message, and whole folder contents
// between PST folders.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MoveItemsToOtherFolders
    {
        public static void Run()
        {
            // Work on a copy: moving items rewrites the PST, and examples must never
            // modify the shared input data.
            var pstPath = Data.Out/"MoveItemsToOtherFolders_out.pst";
            File.Copy(Data.Mapi/"Outlook_1.pst", pstPath, true);

            using (var personalStorage = PersonalStorage.FromFile(pstPath))
            {
                var inbox = personalStorage.GetPredefinedFolder(StandardIpmFolder.Inbox);
                var deleted = personalStorage.GetPredefinedFolder(StandardIpmFolder.DeletedItems);
                var subfolder = inbox.GetSubFolder("SubInbox");

                Console.WriteLine($"Inbox before:         {inbox.ContentCount} item(s), " +
                                  $"{inbox.GetSubFolders().Count} subfolder(s)");
                Console.WriteLine($"Deleted Items before: {deleted.ContentCount} item(s), " +
                                  $"{deleted.GetSubFolders().Count} subfolder(s)");

                // A whole folder, with everything in it.
                personalStorage.MoveItem(subfolder, deleted);

                // Or a single message, addressed by its MessageInfo.
                var contents = inbox.GetContents();
                if (contents.Count > 0)
                    personalStorage.MoveItem(contents[0], deleted);

                // Or everything a folder holds, in one call each: first any remaining
                // subfolders, then the messages.
                inbox.MoveSubfolders(deleted);
                inbox.MoveContents(deleted);

                Console.WriteLine($"\nInbox after:          {inbox.ContentCount} item(s), " +
                                  $"{inbox.GetSubFolders().Count} subfolder(s)");
                Console.WriteLine($"Deleted Items after:  {deleted.ContentCount} item(s), " +
                                  $"{deleted.GetSubFolders().Count} subfolder(s)");
                Console.WriteLine($"Saved to {pstPath}");
            }
        }
    }
}
