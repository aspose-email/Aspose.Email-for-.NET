// Demonstrates creating and updating an Outlook task as a MAPI item, asynchronously.
//
// CreateItemAsync stores a MapiTask in the Tasks folder. FetchItemAsync reads it back as
// a MapiMessage, which ToMapiMessageItem turns into a MapiTask again. After changing
// it, UpdateItemAsync writes it back - EwsUpdateItem.Create has overloads for tasks,
// contacts, notes and calendar items. The example deletes the task at the end.

using System;
using System.Threading.Tasks;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange.WebService.Models;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateMapiTaskEwsAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            using (var client = await ClientBuilder.EwsAsync(AuthType.ModernWithAppPermission))
            {
                var task = new MapiTask("Prepare the quarterly report", "Collect the numbers from all regions.",
                    DateTime.Today, DateTime.Today.AddDays(7));

                var uri = await client.CreateItemAsync(task, client.MailboxInfo.TasksUri);

                try
                {
                    var stored = await FetchTaskAsync(client, uri);
                    Console.WriteLine($"Created: {stored.Subject}, {stored.Status}, {stored.PercentComplete}% done");

                    stored.Status = MapiTaskStatus.InProgress;
                    stored.PercentComplete = 50;
                    await client.UpdateItemAsync(EwsUpdateItem.Create(stored));

                    var updated = await FetchTaskAsync(client, uri);
                    Console.WriteLine($"Updated: {updated.Subject}, {updated.Status}, {updated.PercentComplete}% done");
                }
                finally
                {
                    await client.DeleteItemAsync(uri, DeletionOptions.DeletePermanently);
                }
            }
        }

        private static async Task<MapiTask> FetchTaskAsync(IAsyncEwsClient client, string uri)
        {
            var message = await client.FetchItemAsync(uri);
            return (MapiTask)message.ToMapiMessageItem();
        }
    }
}
