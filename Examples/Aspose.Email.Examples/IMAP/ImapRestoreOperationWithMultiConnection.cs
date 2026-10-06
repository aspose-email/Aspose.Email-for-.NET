// Demonstrates a restore that uploads over several connections at once.
//
// With UseMultiConnection enabled, heavy operations such as Restore spread their work
// over up to ConnectionsQuantity parallel connections, which can shorten the upload of a
// large PST. It does not always help - servers limit and throttle connections per
// account - so measure against your own server.
//
// The example builds a PST with 50 messages in a uniquely named folder, restores it,
// and removes that folder from the mailbox at the end.

using System;
using System.Diagnostics;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapRestoreOperationWithMultiConnection
    {
        public static void Run()
        {
            const int messageCount = 50;

            var folderName = "Aspose-Restore-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var pstPath = Data.Out/"ImapRestoreOperationWithMultiConnection_in.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var folder = pst.RootFolder.AddSubFolder(folderName);
                for (var i = 1; i <= messageCount; i++)
                {
                    folder.AddMessage(new MapiMessage("from@example.com", "to@example.com",
                        $"Restored message {i}", "Created by the ImapRestoreOperationWithMultiConnection example."));
                }
            }

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var pst = PersonalStorage.FromFile(pstPath, false))
            {
                client.UseMultiConnection = MultiConnectionMode.Enable;
                client.ConnectionsQuantity = 5;

                var watch = Stopwatch.StartNew();
                client.Restore(pst, new RestoreSettings { Recursive = true });
                Console.WriteLine($"Restored {messageCount} message(s) over up to {client.ConnectionsQuantity} " +
                                  $"connections in {watch.Elapsed.TotalSeconds:F1} s.");

                try
                {
                    Console.WriteLine(client.ExistFolder(folderName)
                        ? $"'{folderName}' now holds {client.GetFolderInfo(folderName).TotalMessageCount} message(s)."
                        : $"'{folderName}' was not found at the top level of the mailbox.");
                }
                finally
                {
                    if (client.ExistFolder(folderName))
                        client.DeleteFolder(folderName);
                }
            }
        }
    }
}
