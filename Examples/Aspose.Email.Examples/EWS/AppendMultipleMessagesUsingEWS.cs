// Demonstrates uploading several messages to a folder in one request.
//
// AppendMessages takes MailMessage or MapiMessage objects, stores them in the folder and
// returns the item URI of each new message, in the same order. The messages are not
// sent - they simply appear in the folder, which suits importing mail from files or
// another system. The example uses a uniquely named folder and deletes it at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class AppendMultipleMessagesUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 5)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Imported message {i}", "Body"))
                        .ToList();

                    var uris = client.AppendMessages(folder.Uri, messages);

                    Console.WriteLine($"Appended {uris.Length} message(s) to '{folderName}':");
                    foreach (var info in client.ListMessages(folder.Uri))
                        Console.WriteLine($"  {info.Subject}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
