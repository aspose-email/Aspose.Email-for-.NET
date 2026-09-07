// Demonstrates the asynchronous Graph client.
//
// IGraphClientAsync mirrors IGraphClient, with every call returning a Task and taking a
// CancellationToken. That matters against a remote service: the calls are network-bound,
// so awaiting them frees the thread instead of blocking it, and the token lets a slow
// mailbox scan be abandoned.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class UseGraphClientAsync
    {
        // The example runner calls a synchronous Run(), so bridge to the async body here.
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(2)))
            using (var client = ClientBuilder.GraphAsync(AuthType.ModernWithAppPermission))
            {
                var folders = await client.ListFoldersAsync(null, cancellation.Token);
                Console.WriteLine($"{folders.Count} folder(s):");
                foreach (var folder in folders)
                    Console.WriteLine($"  {folder.DisplayName}");

                var inbox = folders.FirstOrDefault(f => f.DisplayName == "Inbox");
                if (inbox == null)
                {
                    Console.WriteLine("No Inbox in this mailbox.");
                    return;
                }

                // The paged overload works the same way as on the synchronous client.
                var page = await client.ListMessagesAsync(inbox.ItemId, new PageInfo(10), null, cancellation.Token);
                Console.WriteLine($"\nFirst page: {page.Items.Count} message(s), last page: {page.LastPage}");

                foreach (MessageInfo messageInfo in page.Items)
                    Console.WriteLine($"  {messageInfo.Subject}");

                var first = page.Items.Cast<MessageInfo>().FirstOrDefault();
                if (first != null)
                {
                    var message = await client.FetchMessageAsync(first.ItemId, cancellation.Token);
                    Console.WriteLine($"\nFetched: {message.Subject} ({message.Attachments.Count} attachment(s))");
                }

                // Several independent calls can be in flight at once.
                var calendarsTask = client.ListCalendarsAsync(null, cancellation.Token);
                var categoriesTask = client.ListCategoriesAsync(null, cancellation.Token);
                var taskListsTask = client.ListTaskListsAsync(null, cancellation.Token);

                await Task.WhenAll(calendarsTask, categoriesTask, taskListsTask);

                Console.WriteLine($"\nCalendars:  {calendarsTask.Result.Count}");
                Console.WriteLine($"Categories: {categoriesTask.Result.Count}");
                Console.WriteLine($"Task lists: {taskListsTask.Result.Count}");
            }
        }
    }
}
