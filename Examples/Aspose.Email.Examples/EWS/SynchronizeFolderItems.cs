// Demonstrates incremental synchronisation of a folder.
//
// The first SyncFolder call returns everything in the folder plus a SyncState. Passing
// that state to the next call returns only what changed since - new, changed and
// deleted items and read-flag changes - which is how a client keeps a local copy up to
// date cheaply. The example works in a uniquely named folder and deletes it at the end.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class SynchronizeFolderItems
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    client.AppendMessage(folder.Uri,
                        new MailMessage("sender@example.com", "receiver@example.com", "Existing message", "Body"));

                    var first = client.SyncFolder(folder.Uri, (string)null);
                    Console.WriteLine($"Initial sync: {first.NewItems.Count} new item(s)");

                    // Changes made after the first sync...
                    client.AppendMessage(folder.Uri,
                        new MailMessage("sender@example.com", "receiver@example.com", "Arrived later", "Body"));

                    // ...are all the next sync returns.
                    var next = client.SyncFolder(first.SyncObject);
                    Console.WriteLine($"Next sync:    {next.NewItems.Count} new, {next.ChangedItems.Count} changed, " +
                                      $"{next.DeletedItems.Length} deleted, " +
                                      $"{next.ReadFlagChanged.Count} read-flag change(s)");
                    foreach (var info in next.NewItems)
                        Console.WriteLine($"  new: {info.Subject}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
