using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class CopyConversations
    {
        public static void Run()
        {   
            var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            // Find those Conversation Items in the Inbox folder which we want to copy
            var conversations = client.FindConversations(client.MailboxInfo.InboxUri);
            
            foreach (var conversation in conversations)
            {
                Console.WriteLine("Topic: " + conversation.ConversationTopic);

                // Copy the conversation item based on some condition
                if (conversation.ConversationTopic.Contains("test email") == true)
                {
                    client.CopyConversationItems(conversation.ConversationId, client.MailboxInfo.DeletedItemsUri);
                    Console.WriteLine("Copied the conversation item to another folder");
                }
            }
        }
    }
}