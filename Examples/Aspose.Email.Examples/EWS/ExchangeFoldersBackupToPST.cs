// Demonstrates backing up Exchange folders to a PST file.
//
// Backup downloads the given folders with their items and writes them into a new
// Outlook PST; BackupOptions.Recursive includes subfolders too. The PST can be opened
// in Outlook, read with PersonalStorage, or uploaded again with Restore. Here only the
// Inbox itself is backed up.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExchangeFoldersBackupToPST
    {
        public static void Run()
        {
            var outputPath = Data.Out/"ExchangeFoldersBackupToPST_out.pst";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inbox = client.GetFolderInfo(client.MailboxInfo.InboxUri);
                Console.WriteLine($"Backing up '{inbox.DisplayName}' ({inbox.TotalCount} item(s))...");

                client.Backup(new ExchangeFolderInfoCollection { inbox }, outputPath, BackupOptions.None);
            }

            using (var pst = PersonalStorage.FromFile(outputPath, false))
            {
                Console.WriteLine($"\nWritten to {outputPath}:");
                foreach (var folder in pst.RootFolder.GetSubFolders())
                    Console.WriteLine($"  {folder.DisplayName}: {folder.ContentCount} item(s)");
            }
        }
    }
}
