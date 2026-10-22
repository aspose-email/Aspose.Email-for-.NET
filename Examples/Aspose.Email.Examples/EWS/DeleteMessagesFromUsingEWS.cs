// Demonstrates deleting messages that match a condition.
//
// DeleteItem takes an item URI and DeletionOptions: MoveToDeletedItems keeps the message
// recoverable, DeletePermanently removes it for good. Here every message whose subject
// contains "delete" is removed. The example works on its own messages in a uniquely
// named folder and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteMessagesFromUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    client.AppendMessages(folder.Uri,
                        new MailMessage("sender@example.com", "receiver@example.com", "Please delete me", "Body"),
                        new MailMessage("sender@example.com", "receiver@example.com", "Keep me", "Body"),
                        new MailMessage("sender@example.com", "receiver@example.com", "Delete this too", "Body"));

                    foreach (var info in client.ListMessages(folder.Uri))
                    {
                        var matches = info.Subject != null &&
                                      info.Subject.IndexOf("delete", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (matches)
                        {
                            client.DeleteItem(info.UniqueUri, DeletionOptions.DeletePermanently);
                            Console.WriteLine($"Deleted: {info.Subject}");
                        }
                    }

                    Console.WriteLine("Left in the folder:");
                    foreach (var info in client.ListMessages(folder.Uri))
                        Console.WriteLine($"  {info.Subject}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
