// Demonstrates backing up mailbox folders to a PST file.
//
// Backup downloads the given folders with their messages and writes them into a new
// Outlook PST; BackupOptions.Recursive includes the subfolders as well. The PST can be
// opened in Outlook, read with PersonalStorage, or uploaded again with Restore - see
// ImapRestoreOperation. Here only the Inbox itself is backed up, to keep the run short.

using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapBackupOperation
    {
        public static void Run()
        {
            var outputPath = Data.Out/"ImapBackupOperation_out.pst";

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var inbox = client.GetFolderInfo(ImapFolderInfo.InBox);
                Console.WriteLine($"Backing up '{inbox.Name}' ({inbox.TotalMessageCount} message(s))...");

                client.Backup(new ImapFolderInfoCollection(inbox), outputPath, BackupOptions.None);
            }

            using (var pst = PersonalStorage.FromFile(outputPath, false))
            {
                Console.WriteLine($"\nWritten to {outputPath}:");
                foreach (var folder in pst.RootFolder.GetSubFolders())
                    Console.WriteLine($"  {folder.DisplayName}: {folder.ContentCount} message(s)");
            }
        }
    }
}
