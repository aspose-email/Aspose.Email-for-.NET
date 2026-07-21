using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class FindConversationsOnExchangeServer
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
            
            // Find Conversation Items in the Inbox folder
            ExchangeConversation[] conversations = client.FindConversations(client.MailboxInfo.InboxUri);
            // Show all conversations
            foreach (ExchangeConversation conversation in conversations)
            {
                // Display conversation properties like Id and Topic
                Console.WriteLine("Topic: " + conversation.ConversationTopic);
                Console.WriteLine("Flag Status: " + conversation.FlagStatus.ToString());
                Console.WriteLine();
            }
        }
    }
}