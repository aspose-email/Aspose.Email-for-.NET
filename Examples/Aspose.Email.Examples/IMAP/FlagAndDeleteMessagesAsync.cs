// Demonstrates changing flags and deleting messages with the task-based API.
//
// ImapChangeMessageFlags is shared by AddMessageFlagsAsync, RemoveMessageFlagsAsync and
// ChangeMessageFlagsAsync. ImapDeleteMessages marks messages as deleted and can expunge
// them right away; here the expunge is done separately with CommitDeletesAsync, limited
// to the given unique ids (this needs UIDPLUS, RFC 4315).

using System;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Imap.Models;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FlagAndDeleteMessagesAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                IAsyncImapClient client = imapClient;

                await client.CreateFolderAsync(folderName);

                try
                {
                    var messages = Enumerable.Range(1, 4)
                        .Select(i => new MailMessage("from@example.com", "to@example.com", $"Alert {i}", "Body"))
                        .ToList();

                    var appended = (AppendMessagesFromMessageObjectResult)await client.AppendMessagesAsync(messages, folderName);
                    var uids = appended.Succeeded.Values.ToList();

                    await client.SelectFolderAsync(folderName);

                    await client.AddMessageFlagsAsync(ImapChangeMessageFlags.Create()
                        .SetMessages(uids)
                        .SetFlags(ImapMessageFlags.IsRead | ImapMessageFlags.Flagged));
                    await PrintAsync(client, folderName, "All read and flagged:");

                    await client.RemoveMessageFlagsAsync(ImapChangeMessageFlags.Create()
                        .SetMessages(uids.Skip(2))
                        .SetFlags(ImapMessageFlags.Flagged));
                    await PrintAsync(client, folderName, "Flag removed from the last two:");

                    // Mark the first two as deleted, then expunge exactly those.
                    var toDelete = uids.Take(2).ToList();
                    await client.DeleteMessagesAsync(ImapDeleteMessages.Create()
                        .SetMessages(toDelete)
                        .SetCommitNow(false));

                    if (imapClient.UidPlusSupported)
                        await client.CommitDeletesAsync(ImapUniqueIdParameterSet.Create().SetMessages(toDelete));
                    else
                        await imapClient.CommitDeletesAsync();

                    await PrintAsync(client, folderName, "After deleting the first two:");
                }
                finally
                {
                    await client.DeleteFolderAsync(folderName);
                }
            }
        }

        private static async Task PrintAsync(IAsyncImapClient client, string folderName, string title)
        {
            Console.WriteLine(title);
            foreach (var info in await client.ListMessagesAsync(folderName))
                Console.WriteLine($"  {info.Subject}: {info.Flags}");
            Console.WriteLine();
        }
    }
}
