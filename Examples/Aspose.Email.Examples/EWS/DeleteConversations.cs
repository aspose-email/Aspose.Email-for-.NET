using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class DeleteConversations
    {
        public static void Run()
        {
            const string mailboxUri = "https://exchange/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);

            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);
            Console.WriteLine("Connected to Exchange 2010");
            
            // Find those Conversation Items in the Inbox folder which we want to delete
            ExchangeConversation[] conversations = client.FindConversations(client.MailboxInfo.InboxUri);
            foreach (ExchangeConversation conversation in conversations)
            {
                Console.WriteLine("Topic: " + conversation.ConversationTopic);
                // Delete the conversation item based on some condition
                if (conversation.ConversationTopic.Contains("test email") == true)
                {
                    client.DeleteConversationItems(conversation.ConversationId);
                    Console.WriteLine("Deleted the conversation item");
                }
            }
        }
    }
}