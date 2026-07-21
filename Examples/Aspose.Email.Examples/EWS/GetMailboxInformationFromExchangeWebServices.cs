using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class GetMailboxInformationFromExchangeWebServices
    {
        public static void Run()
        {
            try
            {
                // Create instance of EWSClient class by giving credentials         
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Get mailbox size, exchange mailbox info, Mailbox and Inbox folder URI
                Console.WriteLine("Mailbox size: " + client.GetMailboxSize() + " bytes");
                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri);
                Console.WriteLine("Inbox folder URI: " + mailboxInfo.InboxUri);
                Console.WriteLine("Sent Items URI: " + mailboxInfo.SentItemsUri);
                Console.WriteLine("Drafts folder URI: " + mailboxInfo.DraftsUri);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }

        }
    }
}