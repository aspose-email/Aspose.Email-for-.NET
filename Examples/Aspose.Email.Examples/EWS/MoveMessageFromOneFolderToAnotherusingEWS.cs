using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class MoveMessageFromOneFolderToAnotherusingEWS
    {
        public static void Run()
        {
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
        }
    }
}