using System;
using System.Net;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class DeleteMessagesFromusingEWS
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
                // Delete message based on some criteria
                if (msgInfo.Subject != null && msgInfo.Subject.ToLower().Contains("delete") == true)
                {
                    client.DeleteItem(msgInfo.UniqueUri, DeletionOptions.DeletePermanently); // EWS
                    Console.WriteLine("Message deleted...." + msgInfo.Subject);
                }
                else
                {
                    // Do something else
                }
            }
        }
    }
}