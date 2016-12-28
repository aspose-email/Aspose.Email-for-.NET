using System;
using System.Net;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
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
            // ExStart:DeleteConversations
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
            // ExEnd:DeleteConversations
        }
    }
}