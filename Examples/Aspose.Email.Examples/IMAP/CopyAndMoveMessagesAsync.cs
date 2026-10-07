// Demonstrates copying and moving messages with the task-based API.
//
// ImapCopyMessages and ImapMoveMessages bundle the parameters of CopyMessagesAsync and
// MoveMessagesAsync: which messages (by unique id, sequence number, range or message
// info), the target folder, and for a move whether the originals are expunged at once.

using System;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Imap.Models;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CopyAndMoveMessagesAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var inboxFolder = "Aspose-Inbox-" + suffix;
            var copiesFolder = "Aspose-Copies-" + suffix;
            var archiveFolder = "Aspose-Archive-" + suffix;

            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                IAsyncImapClient client = imapClient;

                await client.CreateFolderAsync(inboxFolder);
                await client.CreateFolderAsync(copiesFolder);
                await client.CreateFolderAsync(archiveFolder);

                try
                {
                    var messages = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage("from@example.com", "to@example.com", $"Order {i}", "Body"))
                        .ToList();

                    var result = await client.AppendMessagesAsync(messages, inboxFolder);
                    var appended = (AppendMessagesFromMessageObjectResult)result;
                    var uids = appended.Succeeded.Values.ToList();

                    await client.SelectFolderAsync(inboxFolder);

                    // All three into the copies folder.
                    await client.CopyMessagesAsync(ImapCopyMessages.Create()
                        .SetMessages(uids)
                        .SetFolder(copiesFolder));

                    // The first two into the archive.
                    await client.MoveMessagesAsync(ImapMoveMessages.Create()
                        .SetMessages(uids.Take(2))
                        .SetFolder(archiveFolder)
                        .SetCommitDeletions(true));

                    foreach (var folderName in new[] { inboxFolder, copiesFolder, archiveFolder })
                    {
                        var info = await client.GetFolderInfoAsync(folderName);
                        Console.WriteLine($"{folderName}: {info.TotalMessageCount} message(s)");
                    }
                }
                finally
                {
                    await client.DeleteFolderAsync(inboxFolder);
                    await client.DeleteFolderAsync(copiesFolder);
                    await client.DeleteFolderAsync(archiveFolder);
                }
            }
        }
    }
}
