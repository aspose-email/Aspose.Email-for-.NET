// Demonstrates moving a whole conversation to another folder in one call.
//
// MoveConversationItems moves every message of a conversation - only those in the
// context folder, when one is given - which is how "move this thread to Archive" works.
// The example works in two uniquely named folders and deletes both at the end.

using System;
using System.Collections.Specialized;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class MoveConversations
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var source = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Source-" + suffix);
                var target = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Target-" + suffix);

                try
                {
                    client.AppendMessages(source.Uri,
                        new MailMessage("alice@example.com", "bob@example.com", "Offsite agenda", "See you there."));

                    var conversation = client.FindConversations(source.Uri).First();
                    client.MoveConversationItems(conversation.ConversationId, source.Uri, target.Uri);
                    Console.WriteLine($"Moved the conversation '{conversation.ConversationTopic}'.");

                    Report(client, source.DisplayName, source.Uri);
                    Report(client, target.DisplayName, target.Uri);
                }
                finally
                {
                    client.DeleteFolders(new StringCollection { source.Uri, target.Uri }, true);
                }
            }
        }

        private static void Report(IEWSClient client, string name, string folderUri)
        {
            Console.WriteLine($"  '{name}': {client.ListMessages(folderUri).Count} message(s)");
        }
    }
}
