using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class PreFetchMessageSizeUsingIEWSClient
    {
        public static void Run()
        {
            // Create instance of ExchangeWebServiceClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Call ListMessages method to list messages info from Inbox
            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

            // Loop through the collection to display the basic information
            foreach (ExchangeMessageInfo msgInfo in msgCollection)
            {
                Console.WriteLine("Subject: " + msgInfo.Subject);
                Console.WriteLine("From: " + msgInfo.From.ToString());
                Console.WriteLine("To: " + msgInfo.To.ToString());
                Console.WriteLine("Message Size: " + msgInfo.Size);
                Console.WriteLine("==================================");
            }
        }
    }
}