// Demonstrates downloading a message with the task-based API.
//
// FetchMessagesAsync takes an ImapFetchMessages parameter set; SetMessage picks a single
// message by unique id or sequence number. The call hands back a Task, so the caller can
// keep working - here it just reports progress - and await the message when it needs it.

using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Imap.Models;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RetrievingMessagesAsynchronously
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                IAsyncImapClient client = imapClient;

                await client.SelectFolderAsync(ImapFolderInfo.InBox);
                var infos = await client.ListMessagesAsync(ImapFolderInfo.InBox);

                if (infos.Count == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                var newest = infos.OrderByDescending(info => info.InternalDate).First();
                var watch = Stopwatch.StartNew();

                var download = client.FetchMessagesAsync(ImapFetchMessages.Create().SetMessage(newest.UniqueId));
                Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  download of '{newest.Subject}' started");

                while (!download.IsCompleted)
                {
                    Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  still downloading...");
                    await Task.WhenAny(download, Task.Delay(200));
                }

                var message = (await download).First();
                Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  done");

                Console.WriteLine($"\n  subject:     {message.Subject}");
                Console.WriteLine($"  from:        {message.From}");
                Console.WriteLine($"  attachments: {message.Attachments.Count}");
            }
        }
    }
}
