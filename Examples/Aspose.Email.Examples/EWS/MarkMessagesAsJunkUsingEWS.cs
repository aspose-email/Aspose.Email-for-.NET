// Demonstrates reporting messages as junk, and taking that back.
//
// MarkAsJunk(true, ...) adds the senders to the mailbox's blocked senders list, so
// future mail from them goes to Junk Email; with moveItem the messages are moved there
// too. MarkAsJunk(false, ...) removes the senders from the list again. The example works
// on its own test messages in a uniquely named folder, undoes the change, and deletes
// the folder at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class MarkMessagesAsJunkUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var sender = $"spammer-{Guid.NewGuid():N}@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 2)
                        .Select(i => new MailMessage(sender, "receiver@example.com", $"You have won {i}", "Body"));
                    var uris = client.AppendMessages(folder.Uri, messages);

                    // Block the sender, but leave the messages where they are.
                    var marked = client.MarkAsJunk(true, false, uris);
                    Console.WriteLine($"Reported {marked.Length} message(s) as junk; {sender} is now blocked.");

                    // Unblock the sender again.
                    client.MarkAsJunk(false, false, uris);
                    Console.WriteLine($"Reported them as not junk; {sender} is no longer blocked.");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
