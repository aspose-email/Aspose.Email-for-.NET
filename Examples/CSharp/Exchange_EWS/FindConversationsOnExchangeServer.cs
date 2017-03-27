using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
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
            
            // ExStart:FindConversationsOnExchangeServer
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
            // ExEnd:FindConversationsOnExchangeServer
        }
    }
}