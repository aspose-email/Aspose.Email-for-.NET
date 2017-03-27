using System;
using System.Net;
using Aspose.Email.Mime;
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
    class MoveMessageFromOneFolderToAnotherusingEWS
    {
        public static void Run()
        {
            // ExStart:MoveMessageFromOneFolderToAnotherusingEWS
            // Create instance of IEWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();

            // List all messages from Inbox folder
            Console.WriteLine("Listing all messages from Inbox....");
            ExchangeMessageInfoCollection msgInfoColl = client.ListMessages(mailboxInfo.InboxUri);
            foreach (ExchangeMessageInfo msgInfo in msgInfoColl)
            {
                // Move message to "Processed" folder, after processing certain messages based on some criteria
                if (msgInfo.Subject != null &&
                    msgInfo.Subject.ToLower().Contains("process this message") == true)
                {
                    client.MoveItem(mailboxInfo.DeletedItemsUri, msgInfo.UniqueUri); // EWS
                    Console.WriteLine("Message moved...." + msgInfo.Subject);
                }
                else
                {
                    // Do something else
                }
            }
            // ExEnd:MoveMessageFromOneFolderToAnotherusingEWS
        }
    }
}