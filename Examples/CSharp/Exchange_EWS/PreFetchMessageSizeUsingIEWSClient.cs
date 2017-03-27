using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class PreFetchMessageSizeUsingIEWSClient
    {
        public static void Run()
        {
            // ExStart:PreFetchMessageSizeUsingIEWSClient
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
            // ExEnd:PreFetchMessageSizeUsingIEWSClient
        }
    }
}