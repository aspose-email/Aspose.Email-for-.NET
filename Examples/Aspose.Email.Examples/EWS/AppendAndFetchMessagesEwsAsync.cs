// Demonstrates bulk operations of the task-based EWS client.
//
// The bulk methods take a parameter object from the
// Aspose.Email.Clients.Exchange.WebService.Models namespace: EwsAppendMessage collects
// the messages and the target folder for AppendMessagesAsync, EwsFetchItems the item
// URIs for FetchItemsAsync - with SkipAttachments to leave attachment content on the
// server. Both accept a cancellation token. The example works in a uniquely named
// folder and deletes it at the end.

using System;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange.WebService.Models;

namespace Aspose.Email.Examples.EWS
{
    internal static class AppendAndFetchMessagesEwsAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = await ClientBuilder.EwsAsync(AuthType.ModernWithAppPermission))
            {
                var folder = await client.CreateFolderAsync(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Report {i}", "Body"));

                    var uris = (await client.AppendMessagesAsync(EwsAppendMessage.Create()
                        .SetFolder(folder.Uri)
                        .AddMessages(messages))).ToList();
                    Console.WriteLine($"Appended {uris.Count} message(s) to '{folderName}'.");

                    var fetched = await client.FetchItemsAsync(EwsFetchItems.Create()
                        .AddUris(uris)
                        .SkipAttachments());

                    Console.WriteLine("Fetched back as MapiMessage objects:");
                    foreach (var message in fetched)
                        Console.WriteLine($"  {message.Subject} ({message.MessageClass})");

                    await client.MarkAllItemsAsync(new[] { folder.Uri }, read: true);
                    await client.DeleteItemsAsync(uris.Take(2), DeletionOptions.DeletePermanently);

                    var remaining = await client.ListMessagesAsync(folder.Uri);
                    Console.WriteLine($"\nAfter marking all read and deleting two: {remaining.Count} message(s), " +
                                      $"{remaining.Count(info => info.IsRead)} read");
                }
                finally
                {
                    await client.DeleteFolderAsync(folder.Uri, true);
                }
            }
        }
    }
}
