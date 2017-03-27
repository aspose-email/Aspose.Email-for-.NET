using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class MoveMessageFromOneFolderToAnotherUsingExchangeClient
    {
        public static void Run()
        {
            // ExStart:MoveMessageFromOneFolderToAnotherUsingExchangeClient
            string mailboxURI = "https://Ex2003/exchange/administrator"; // WebDAV

            string username = "administrator";
            string password = "pwd";
            string domain = "domain.local";

            Console.WriteLine("Connecting to Exchange Server....");
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            ExchangeClient client = new ExchangeClient(mailboxURI, credential);

            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();

            // List all messages from Inbox folder
            Console.WriteLine("Listing all messages from Inbox....");
            ExchangeMessageInfoCollection msgInfoColl = client.ListMessages(mailboxInfo.InboxUri);
            foreach (ExchangeMessageInfo msgInfo in msgInfoColl)
            {
                // Nove message to "Processed" folder, after processing certain messages based on some criteria
                if (msgInfo.Subject != null &&
                    msgInfo.Subject.ToLower().Contains("process this message") == true)
                {                    
                    client.MoveItems(msgInfo.UniqueUri, client.MailboxInfo.RootUri + "/Processed/" + msgInfo.Subject);
                    Console.WriteLine("Message moved...." + msgInfo.Subject);
                }
                else
                {
                    // Do something else
                }
            }
            // ExEnd:MoveMessageFromOneFolderToAnotherUsingExchangeClient
        }
    }
}