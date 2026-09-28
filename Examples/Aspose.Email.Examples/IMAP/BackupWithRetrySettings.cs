// Demonstrates a folder backup to PST that rides out an unreliable connection.
//
// BackupSettings is the richer counterpart of the BackupOptions flags used by
// ImapBackupOperation. Besides recursion it tells the client to reconnect when the
// server drops the connection, and how many times to repeat a failed command with what
// pause in between - worthwhile for large mailboxes and for servers that throttle long
// downloads. Here the backup is written to a stream.

using System;
using System.IO;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class BackupWithRetrySettings
    {
        public static void Run()
        {
            var outputPath = Data.Out/"BackupWithRetrySettings_out.pst";

            var settings = new BackupSettings
            {
                ExecuteRecursively = false,     // the Inbox only, not its subfolders
                RestoreConnection = true,       // reconnect if the server hangs up
                NumberOfAttemptsToRrepeat = 3,  // retry a failed command up to 3 times
                TimeoutBetweenAttempts = 5000   // milliseconds between retries
            };

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var inbox = client.GetFolderInfo(ImapFolderInfo.InBox);
                Console.WriteLine($"Backing up '{inbox.Name}' ({inbox.TotalMessageCount} message(s))...");

                using (var stream = File.Create(outputPath))
                    client.Backup(new ImapFolderInfoCollection(inbox), stream, settings);
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
