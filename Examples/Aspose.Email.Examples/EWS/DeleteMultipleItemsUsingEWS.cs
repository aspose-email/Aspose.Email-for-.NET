// Demonstrates deleting several items in one request.
//
// DeleteItems takes a list of item URIs and DeletionOptions: MoveToDeletedItems keeps
// them recoverable, DeletePermanently removes them for good. One request instead of one
// per item matters when cleaning up many messages. The example uses a uniquely named
// folder and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteMultipleItemsUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 5)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Message {i}", "Body"));
                    var uris = client.AppendMessages(folder.Uri, messages);

                    Console.WriteLine($"Before: {client.ListMessages(folder.Uri).Count} message(s)");

                    // Delete the first three for good.
                    client.DeleteItems(uris.Take(3), DeletionOptions.DeletePermanently);

                    Console.WriteLine($"After:  {client.ListMessages(folder.Uri).Count} message(s)");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
