// Demonstrates working with a conversation as a whole.
//
// FindConversations lists the conversations of a folder; FetchConversationMessages
// downloads every message of one conversation, and SetConversationReadState marks all
// of them as read or unread in one call - what Outlook does in conversation view.
// The example works on its own messages in a uniquely named folder and deletes it at
// the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ReadConversationMessages
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
                        new MailMessage("alice@example.com", "bob@example.com", "Project status", "Where are we?"),
                        new MailMessage("bob@example.com", "alice@example.com", "RE: Project status", "On track."));

                    var conversations = client.FindConversations(folder.Uri);
                    Console.WriteLine($"{conversations.Length} conversation(s) in '{folderName}':");
                    foreach (var conversation in conversations)
                    {
                        Console.WriteLine($"  {conversation.ConversationTopic}: " +
                                          $"{conversation.MessageCount} message(s), {conversation.UnreadCount} unread");
                    }

                    var first = conversations.FirstOrDefault();
                    if (first == null)
                        return;

                    var messages = client.FetchConversationMessages(first.ConversationId);
                    Console.WriteLine($"\nMessages of '{first.ConversationTopic}':");
                    foreach (var message in messages)
                        Console.WriteLine($"  {message.From}: {message.Subject}");

                    client.SetConversationReadState(first.ConversationId, true);
                    var after = client.FindConversations(folder.Uri)
                        .First(c => c.ConversationId == first.ConversationId);
                    Console.WriteLine($"\nAfter marking it read: {after.UnreadCount} unread");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
