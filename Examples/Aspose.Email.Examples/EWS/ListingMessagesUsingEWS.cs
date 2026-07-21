using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class ListingMessagesUsingEWS
    {
        public static void Run()
        {
            try
            {
                // Create instance of ExchangeWebServiceClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "UserName", "Password");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to display the basic information
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    Console.WriteLine("Subject: " + msgInfo.Subject);
                    Console.WriteLine("From: " + msgInfo.From.ToString());
                    Console.WriteLine("To: " + msgInfo.To.ToString());
                    Console.WriteLine("Message ID: " + msgInfo.MessageId);
                    Console.WriteLine("Unique URI: " + msgInfo.UniqueUri);
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}