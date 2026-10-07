// Demonstrates uploading a batch of messages over several connections at once.
//
// AppendMessages uploads a whole list and reports, per message, the unique id it got or
// the error that stopped it. With UseMultiConnection enabled, the uploads are spread
// over up to ConnectionsQuantity parallel connections.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapGroupAppendWithMultiConnection
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    client.UseMultiConnection = MultiConnectionMode.Enable;
                    client.ConnectionsQuantity = 5;

                    var messages = Enumerable.Range(1, 20)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Batch message {i}", "Appended in a group over several connections."))
                        .ToList();

                    var result = (AppendMessagesFromMessageObjectResult)client.AppendMessages(folderName, messages);

                    Console.WriteLine($"Appended:    {result.Succeeded.Count}");
                    Console.WriteLine($"Failed:      {result.Failed.Count}");
                    Console.WriteLine($"Not handled: {result.NotHandled.Count}");
                    var total = client.GetFolderInfo(folderName).TotalMessageCount;
                    Console.WriteLine($"\n'{folderName}' now holds {total} message(s).");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
