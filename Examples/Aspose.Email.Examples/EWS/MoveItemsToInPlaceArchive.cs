// Demonstrates moving items to the mailbox's In-Place Archive (online archive).
//
// ArchiveItem moves an item from the given folder to the archive mailbox, keeping its
// folder structure there - the same as Outlook's "Archive" for archive-enabled
// mailboxes. The mailbox needs an archive. The example archives two test messages from
// a uniquely named folder; they stay in the archive afterwards, so delete them there
// when you are done.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class MoveItemsToInPlaceArchive
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 2)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Archive me {i}", "Body"));
                    var uris = client.AppendMessages(folder.Uri, messages);

                    foreach (var uri in uris)
                        client.ArchiveItem(folder.Uri, uri);

                    var left = client.ListMessages(folder.Uri).Count;
                    Console.WriteLine($"Archived {uris.Length} message(s); {left} left in '{folderName}'.");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
