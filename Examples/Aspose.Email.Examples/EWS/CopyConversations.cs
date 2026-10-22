// Demonstrates copying a whole conversation to another folder in one call.
//
// CopyConversationItems copies every message of a conversation; the context folder
// limits it to the messages in that folder. The originals stay where they are.
// The example works in two uniquely named folders and deletes both at the end.

using System;
using System.Collections.Specialized;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class CopyConversations
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
                        new MailMessage("alice@example.com", "bob@example.com", "Budget 2027", "Draft attached."));

                    var conversation = client.FindConversations(source.Uri).First();
                    client.CopyConversationItems(conversation.ConversationId, source.Uri, target.Uri);
                    Console.WriteLine($"Copied the conversation '{conversation.ConversationTopic}'.");

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
