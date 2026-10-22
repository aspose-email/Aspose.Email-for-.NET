// Demonstrates deleting a whole conversation in one call.
//
// DeleteConversationItems removes every message of a conversation - only those in the
// context folder, when one is given - like "Delete conversation" in Outlook.
// The example deletes a conversation it created in a uniquely named folder and deletes
// the folder at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteConversations
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    client.AppendMessages(folder.Uri,
                        new MailMessage("alice@example.com", "bob@example.com", "Old thread", "Body"),
                        new MailMessage("carol@example.com", "bob@example.com", "Keep this", "Body"));

                    Console.WriteLine($"Before: {client.ListMessages(folder.Uri).Count} message(s)");

                    var conversation = client.FindConversations(folder.Uri)
                        .First(c => c.ConversationTopic == "Old thread");
                    client.DeleteConversationItems(conversation.ConversationId, folder.Uri);
                    Console.WriteLine($"Deleted the conversation '{conversation.ConversationTopic}'.");

                    Console.WriteLine($"After:  {client.ListMessages(folder.Uri).Count} message(s)");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
