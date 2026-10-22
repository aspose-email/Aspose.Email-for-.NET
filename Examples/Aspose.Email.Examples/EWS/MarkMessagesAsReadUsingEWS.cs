// Demonstrates changing the read state of messages.
//
// SetReadFlag marks one message as read or unread. MarkAllItems changes every item in
// the given folders in one request; suppressReadReceipts keeps Exchange from sending
// read receipts that the senders may have requested. The example uses a uniquely named
// folder and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class MarkMessagesAsReadUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Message {i}", "Body"));
                    var uris = client.AppendMessages(folder.Uri, messages);

                    client.SetReadFlag(uris[0], true);
                    Report(client, folder.Uri, "After SetReadFlag on the first:");

                    client.MarkAllItems(true, true, folder.Uri);
                    Report(client, folder.Uri, "After MarkAllItems(read):");

                    client.MarkAllItems(false, true, folder.Uri);
                    Report(client, folder.Uri, "After MarkAllItems(unread):");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }

        private static void Report(IEWSClient client, string folderUri, string title)
        {
            Console.WriteLine(title);
            foreach (var info in client.ListMessages(folderUri))
                Console.WriteLine($"  {info.Subject}: {(info.IsRead ? "read" : "unread")}");
        }
    }
}
