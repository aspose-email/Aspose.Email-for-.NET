// Demonstrates how to iterate through a PST folder and delete messages matching a condition.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DeleteMessagesFromPstFiles
    {
        public static void Run()
        {
            // Work on a copy: deleting is destructive, and examples must never modify
            // the shared input data.
            var pstPath = Data.Out/"DeleteMessagesFromPstFiles_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var folder = pst.GetPredefinedFolder(StandardIpmFolder.SentItems);
                var deleted = 0;

                // GetContents() is materialised before deleting, because removing items
                // while enumerating the folder would invalidate the enumeration.
                foreach (var msg in folder.GetContents())
                {
                    Console.WriteLine($"{msg.Subject}: {msg.EntryIdString}");

                    // Replace this with your own condition - this one simply matches
                    // messages that the sample PST is known to contain.
                    if (msg.Subject == "message 2")
                    {
                        folder.DeleteChildItem(msg.EntryId);
                        Console.WriteLine("    -> deleted");
                        deleted++;
                    }
                }

                Console.WriteLine($"\nDeleted {deleted} message(s) from {pstPath}");
            }
        }
    }
}
