// Demonstrates how to query a PST folder and delete all matching messages in a single
// bulk operation, which is much faster than deleting them one by one.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DeleteBulkItemsFromPstFile
    {
        public static void Run()
        {
            // Work on a copy: deleting is destructive, and examples must never modify
            // the shared input data.
            var pstPath = Data.Out/"DeleteBulkItemsFromPstFile_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");
                Console.WriteLine($"Messages in Inbox before: {inbox.GetContents().Count}");

                // Replace this with your own criteria - this one simply matches messages
                // that the sample PST is known to contain.
                var queryBuilder = new PersonalStorageQueryBuilder();
                queryBuilder.Subject.Contains("Test Message");

                var deleteList = new List<string>();
                foreach (var msg in inbox.GetContents(queryBuilder.GetQuery()))
                    deleteList.Add(msg.EntryIdString);

                // One call for the whole batch, rather than a DeleteChildItem per message.
                inbox.DeleteChildItems(deleteList);

                Console.WriteLine($"Deleted {deleteList.Count} matching message(s).");
                Console.WriteLine($"Messages in Inbox after:  {inbox.GetContents().Count}");
                Console.WriteLine($"\nSaved to {pstPath}");
            }
        }
    }
}
