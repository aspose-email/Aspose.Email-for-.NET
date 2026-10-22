// Demonstrates the task-based EWS client.
//
// IAsyncEwsClient is a separate interface from IEWSClient: EWSClient.GetEwsClientAsync
// creates it (ClientBuilder.EwsAsync wraps that call with the settings from
// clientsettings.json). Its methods return tasks and accept a CancellationToken, so a
// UI or a service can stay responsive and give up on slow requests.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Email.Examples.EWS
{
    internal static class ConnectWithAsyncEwsClient
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            using (var client = await ClientBuilder.EwsAsync(AuthType.ModernWithAppPermission))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(1)))
            {
                try
                {
                    var mailbox = await client.GetMailboxInfoAsync(cancellationToken: cancellation.Token);
                    Console.WriteLine($"Inbox: {mailbox.InboxUri}");

                    var inbox = await client.GetFolderInfoAsync(mailbox.InboxUri, cancellation.Token);
                    Console.WriteLine($"{inbox.DisplayName}: {inbox.TotalCount} item(s), {inbox.UnreadCount} unread, " +
                                      $"{inbox.ChildFolderCount} subfolder(s)");

                    var messages = await client.ListMessagesAsync(mailbox.InboxUri, maxNumberOfMessages: 5,
                        cancellationToken: cancellation.Token);

                    Console.WriteLine($"\nFirst {messages.Count} message(s) in the Inbox:");
                    foreach (var info in messages.OrderByDescending(info => info.InternalDate))
                        Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Gave up: the requests took longer than a minute.");
                }
            }
        }
    }
}
