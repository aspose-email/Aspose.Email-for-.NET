// Demonstrates copying items as full-fidelity data streams (EWS ExportItems and
// UploadItems).
//
// ExportItems returns each item as an opaque ExchangeStreamedItem that keeps every
// property Exchange stores. CreateItems uploads such streams into a folder - of the same
// mailbox or of another one - which makes the pair a lossless way to migrate or back up
// individual items. UpdateItems overwrites existing items with exported data instead.
// The example copies three Inbox messages into a uniquely named folder and deletes the
// folder at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExportAndImportItemsUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var ids = client.ListMessages(client.MailboxInfo.InboxUri, 3)
                    .Select(info => info.UniqueUri)
                    .ToArray();

                if (ids.Length == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                var exported = client.ExportItems(ids);
                var totalBytes = exported.Sum(item => item.Data.Length);
                Console.WriteLine($"Exported {exported.Length} item(s), {totalBytes} bytes in total.");

                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var results = client.CreateItems(exported, folder.Uri);
                    foreach (var result in results)
                    {
                        var outcome = result.IsSuccessful ? "succeeded" : "failed: " + result.ErrorMessage;
                        Console.WriteLine($"  upload {outcome}");
                    }

                    Console.WriteLine($"\n'{folderName}' now holds:");
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
