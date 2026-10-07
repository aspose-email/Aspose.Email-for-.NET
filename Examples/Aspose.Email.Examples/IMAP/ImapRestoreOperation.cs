// Demonstrates how to upload the contents of a PST file to the mailbox.
//
// Restore recreates the folders of the PST on the server and appends their messages;
// with RestoreSettings.Recursive the subfolders are restored as well. A PST written by
// ImapBackupOperation can be restored the same way. To stay self-contained, the example
// first builds a small PST - one uniquely named folder with a subfolder - and removes
// the restored folders from the mailbox at the end.

using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapRestoreOperation
    {
        public static void Run()
        {
            var folderName = "Aspose-Restore-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var pstPath = Data.Out/"ImapRestoreOperation_in.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var folder = pst.RootFolder.AddSubFolder(folderName);
                folder.AddMessage(new MapiMessage("from@example.com", "to@example.com", "Project plan", "Body"));

                var subfolder = folder.AddSubFolder("Reports");
                subfolder.AddMessage(new MapiMessage("from@example.com", "to@example.com", "Weekly report", "Body"));
            }

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var pst = PersonalStorage.FromFile(pstPath, false))
            {
                var settings = new RestoreSettings { Recursive = true };
                client.Restore(pst, settings);

                try
                {
                    if (!client.ExistFolder(folderName))
                    {
                        Console.WriteLine($"'{folderName}' was not found at the top level of the mailbox.");
                        return;
                    }

                    Console.WriteLine("Restored:");
                    PrintCount(client, folderName);
                    foreach (var child in client.ListFolders(folderName))
                        PrintCount(client, child.Name);
                }
                finally
                {
                    if (client.ExistFolder(folderName))
                    {
                        foreach (var child in client.ListFolders(folderName))
                            client.DeleteFolder(child.Name);
                        client.DeleteFolder(folderName);
                    }
                }
            }
        }

        private static void PrintCount(ImapClient client, string folderName)
        {
            Console.WriteLine($"  {folderName}: {client.GetFolderInfo(folderName).TotalMessageCount} message(s)");
        }
    }
}
