// Demonstrates how to delete a single item from a PST by its entry id, without
// having to locate the folder that holds it first.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DeleteItemFromPST
    {
        public static void Run()
        {
            // Work on a copy: deleting is destructive.
            var pstPath = Data.Out/"DeleteItemFromPST_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");
                Console.WriteLine($"Messages in Inbox before: {inbox.GetContents().Count}");

                var messageInfo = inbox.GetContents()[0];
                Console.WriteLine($"Deleting: {messageInfo.Subject}");

                // DeleteItem takes an entry id, so it works for messages and folders alike.
                pst.DeleteItem(messageInfo.EntryIdString);

                Console.WriteLine($"Messages in Inbox after:  {pst.RootFolder.GetSubFolder("Inbox").GetContents().Count}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}
