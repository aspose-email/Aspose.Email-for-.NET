// Demonstrates listing messages with the task-based API.
//
// IAsyncImapClient.ListMessagesAsync comes in two shapes: one lists a folder (optionally
// only the messages changed since a MODSEQ, or recursively with subfolders), the other
// lists the messages that match a search query. Both accept a CancellationToken, so a
// slow listing of a large folder can be abandoned. ImapClient implements
// IAsyncImapClient, so any client can be used this way.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMessagesAsynchronously
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

                    var all = await client.ListMessagesAsync(ImapFolderInfo.InBox, token: cancellation.Token);
                    Console.WriteLine($"The Inbox holds {all.Count} message(s).");

                    var builder = new ImapQueryBuilder();
                    builder.Subject.Contains("invoice");
                    builder.InternalDate.Since(DateTime.Today.AddDays(-90));

                    var matches = await client.ListMessagesAsync(builder.GetQuery(), ImapFolderInfo.InBox,
                        token: cancellation.Token);

                    Console.WriteLine($"\n'invoice' in the subject, last 90 days: {matches.Count} message(s)");
                    foreach (var info in matches.OrderByDescending(info => info.InternalDate).Take(10))
                        Console.WriteLine($"  {info.InternalDate:d}  {info.Subject}");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Gave up: listing took longer than a minute.");
                }
            }
        }
    }
}
