// Demonstrates a backup that downloads over several connections at once.
//
// With UseMultiConnection enabled, Backup spreads the downloads over up to
// ConnectionsQuantity parallel connections, which can shorten the backup of a large
// folder. Servers limit and throttle connections per account, so more is not always
// faster - measure against your own server.

using System;
using System.Diagnostics;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapBackupOperationWithMultiConnection
    {
        public static void Run()
        {
            var outputPath = Data.Out/"ImapBackupOperationWithMultiConnection_out.pst";

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.UseMultiConnection = MultiConnectionMode.Enable;
                client.ConnectionsQuantity = 5;

                var inbox = client.GetFolderInfo(ImapFolderInfo.InBox);
                Console.WriteLine($"Backing up '{inbox.Name}' ({inbox.TotalMessageCount} message(s)) " +
                                  $"over up to {client.ConnectionsQuantity} connections...");

                var watch = Stopwatch.StartNew();
                client.Backup(new ImapFolderInfoCollection(inbox), outputPath, BackupOptions.None);
                Console.WriteLine($"Done in {watch.Elapsed.TotalSeconds:F1} s.");
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
