// Demonstrates fetching several messages with the task-based API.
//
// The bulk methods of IAsyncImapClient take a parameter object from the
// Aspose.Email.Clients.Imap.Models namespace instead of a long list of overloads:
// ImapFetchMessages.Create() starts one, SetMessages chooses the messages and
// SetCancellationToken lets the caller abandon a download that takes too long.
// ImapClient implements IAsyncImapClient, so any client can be used this way.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Imap.Models;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FetchMessagesAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(1)))
            {
                IAsyncImapClient client = imapClient;

                try
                {
                    await client.SelectFolderAsync(ImapFolderInfo.InBox, token: cancellation.Token);
                    var infos = await client.ListMessagesAsync(ImapFolderInfo.InBox, token: cancellation.Token);

                    var newest = infos.OrderByDescending(info => info.InternalDate).Take(5).ToList();

                    var parameters = ImapFetchMessages.Create()
                        .SetMessages(newest)
                        .SetCancellationToken(cancellation.Token);

                    var messages = await client.FetchMessagesAsync(parameters);

                    Console.WriteLine($"Fetched the {newest.Count} newest message(s):");
                    foreach (var message in messages)
                        Console.WriteLine($"  {message.Date:g}  {message.Subject} ({message.Attachments.Count} attachment(s))");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Gave up: fetching took longer than a minute.");
                }
            }
        }
    }
}
